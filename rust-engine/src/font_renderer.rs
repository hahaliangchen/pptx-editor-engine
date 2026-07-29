use crate::text_layout::{is_vertical_rotated_char, FontMetricSample};
use cosmic_text::{Buffer, Color, FontSystem, SwashCache};

pub fn blend_pixel(target: &mut [u8], index: usize, color: Color) {
    let source_alpha = color.a() as f32 / 255.0;
    if source_alpha <= 0.0 {
        return;
    }
    let destination_alpha = target[index + 3] as f32 / 255.0;
    let output_alpha = source_alpha + destination_alpha * (1.0 - source_alpha);
    if output_alpha <= 0.0 {
        return;
    }
    for (offset, source) in [color.r(), color.g(), color.b()].iter().enumerate() {
        let destination = target[index + offset] as f32;
        target[index + offset] = (((*source as f32 * source_alpha)
            + destination * destination_alpha * (1.0 - source_alpha))
            / output_alpha)
            .round() as u8;
    }
    target[index + 3] = (output_alpha * 255.0).round() as u8;
}

pub fn alpha_vertical_bounds_in_range(
    source: &[u8],
    width: u32,
    height: u32,
    start_row: u32,
    end_row: u32,
) -> Option<(u32, u32)> {
    let width = width as usize;
    let height = height as usize;
    if width == 0 || height == 0 || source.len() < width * height * 4 {
        return None;
    }

    let start_row = (start_row as usize).min(height);
    let end_row = (end_row as usize).min(height).max(start_row);
    let mut top = end_row;
    let mut bottom = 0_usize;
    for y in start_row..end_row {
        let row_start = y * width * 4;
        let row_has_alpha = (0..width).any(|x| source[row_start + x * 4 + 3] != 0);
        if row_has_alpha {
            top = top.min(y);
            bottom = bottom.max(y + 1);
        }
    }

    (top < bottom).then_some((top as u32, bottom as u32))
}

pub fn build_reflection_bitmap(
    source: &[u8],
    width: u32,
    source_height: u32,
    reflection_height: u32,
    scale_y: f32,
    start_alpha: f32,
    end_alpha: f32,
    end_position: f32,
    blur_radius: f32,
) -> Vec<u8> {
    let width = width as usize;
    let source_height = source_height as usize;
    let reflection_height = reflection_height as usize;
    let mut reflected = vec![0_u8; width * reflection_height * 4];
    let scale_y = scale_y.abs().max(0.01);
    // `reflection_height` is the visible portion selected by endPos. The
    // source glyph must still be sampled against the full reflected height;
    // using the cropped height here mirrors only the upper part of the glyph
    // whenever endPos is below 100%.
    let full_reflection_height = (source_height as f32 * scale_y).ceil().max(1.0) as usize;
    let end_position = end_position.clamp(0.0, 1.0);

    for target_y in 0..reflection_height {
        let source_y = ((full_reflection_height
            .saturating_sub(1)
            .saturating_sub(target_y.min(full_reflection_height.saturating_sub(1)))
            as f32
            / scale_y)
            .floor() as usize)
            .min(source_height.saturating_sub(1));
        let position = if full_reflection_height <= 1 {
            1.0
        } else {
            target_y as f32 / (full_reflection_height - 1) as f32
        };
        let fade = if end_position > 0.0 {
            (position / end_position).clamp(0.0, 1.0)
        } else {
            1.0
        };
        let opacity = (start_alpha + (end_alpha - start_alpha) * fade).clamp(0.0, 1.0);

        for x in 0..width {
            let source_index = (source_y * width + x) * 4;
            let target_index = (target_y * width + x) * 4;
            reflected[target_index] = source[source_index];
            reflected[target_index + 1] = source[source_index + 1];
            reflected[target_index + 2] = source[source_index + 2];
            reflected[target_index + 3] = (source[source_index + 3] as f32 * opacity).round() as u8;
        }
    }

    let radius = (blur_radius * 0.5).ceil().clamp(0.0, 8.0) as i32;
    if radius == 0 {
        return reflected;
    }

    let mut blurred = reflected.clone();
    for y in 0..reflection_height {
        for x in 0..width {
            let mut sum = 0_u32;
            let mut count = 0_u32;
            for offset in -radius..=radius {
                let sample_x = (x as i32 + offset).clamp(0, width as i32 - 1) as usize;
                sum += reflected[(y * width + sample_x) * 4 + 3] as u32;
                count += 1;
            }
            blurred[(y * width + x) * 4 + 3] = (sum / count) as u8;
        }
    }
    blurred
}

/// Builds a small run-level shadow bitmap from the already rasterized glyph
/// alpha. This keeps the shadow tied to the shaped glyphs instead of applying
/// one CSS shadow to the whole text box.
pub fn build_shadow_bitmap(
    source: &[u8],
    width: u32,
    height: u32,
    color: Color,
    opacity: f32,
    blur_radius: f32,
    offset_x: f32,
    offset_y: f32,
) -> Vec<u8> {
    let width = width as usize;
    let height = height as usize;
    let mut shadow = vec![0_u8; width * height * 4];
    if width == 0 || height == 0 || source.len() < width * height * 4 {
        return shadow;
    }
    let opacity = opacity.clamp(0.0, 1.0);
    let red = color.r();
    let green = color.g();
    let blue = color.b();

    // DrawingML's blurRad is a blur radius, not a square-kernel width. A
    // separable Gaussian is a much closer match to the Office/WPS shadow
    // profile than the old clipped box average, and retaining f32 offsets
    // avoids moving diagonal shadows by a whole raster pixel.
    let radius = blur_radius.ceil().clamp(0.0, 24.0) as i32;
    let sigma = (blur_radius * 0.5).max(0.01);
    let mut weights = Vec::with_capacity((radius * 2 + 1) as usize);
    let mut weight_sum = 0.0_f32;
    for delta in -radius..=radius {
        let value = (-(delta as f32).powi(2) / (2.0 * sigma * sigma)).exp();
        weights.push(value);
        weight_sum += value;
    }
    if weight_sum > 0.0 {
        for weight in &mut weights {
            *weight /= weight_sum;
        }
    }

    let mut translated = vec![0.0_f32; width * height];
    for y in 0..height {
        for x in 0..width {
            let source_x = x as f32 - offset_x;
            let source_y = y as f32 - offset_y;
            if source_x < 0.0
                || source_y < 0.0
                || source_x > (width - 1) as f32
                || source_y > (height - 1) as f32
            {
                continue;
            }
            let x0 = source_x.floor() as usize;
            let y0 = source_y.floor() as usize;
            let x1 = (x0 + 1).min(width - 1);
            let y1 = (y0 + 1).min(height - 1);
            let fx = source_x - x0 as f32;
            let fy = source_y - y0 as f32;
            let alpha = |sx: usize, sy: usize| source[(sy * width + sx) * 4 + 3] as f32;
            let top = alpha(x0, y0) * (1.0 - fx) + alpha(x1, y0) * fx;
            let bottom = alpha(x0, y1) * (1.0 - fx) + alpha(x1, y1) * fx;
            translated[y * width + x] = (top * (1.0 - fy) + bottom * fy) * opacity;
        }
    }

    let mut horizontal = vec![0.0_f32; width * height];
    for y in 0..height {
        for x in 0..width {
            let mut value = 0.0_f32;
            for delta in -radius..=radius {
                let sample_x = (x as i32 + delta).clamp(0, width as i32 - 1) as usize;
                value += translated[y * width + sample_x] * weights[(delta + radius) as usize];
            }
            horizontal[y * width + x] = value;
        }
    }

    for y in 0..height {
        for x in 0..width {
            let mut value = 0.0_f32;
            for delta in -radius..=radius {
                let sample_y = (y as i32 + delta).clamp(0, height as i32 - 1) as usize;
                value += horizontal[sample_y * width + x] * weights[(delta + radius) as usize];
            }
            let alpha = value.round().clamp(0.0, 255.0) as u8;
            if alpha == 0 {
                continue;
            }
            let index = (y * width + x) * 4;
            shadow[index] = red;
            shadow[index + 1] = green;
            shadow[index + 2] = blue;
            shadow[index + 3] = alpha;
        }
    }
    shadow
}

#[allow(clippy::too_many_arguments)]
pub fn rasterize_buffer(
    font_system: &mut FontSystem,
    swash_cache: &mut SwashCache,
    buffer: &Buffer,
    origin_x: f32,
    origin_y: f32,
    device_scale: f32,
    pixels: &mut [u8],
    bitmap_width: u32,
    bitmap_height: u32,
    first_line_offset: f32,
    subsequent_line_offset: f32,
    hanging_punctuation: bool,
    vertical_text: bool,
    vertical_line_offsets: &[f32],
    vertical_column_height: f32,
    column_width: f32,
    font_alignment: &str,
    font_size: f32,
    metric_line_height: f32,
    requested_line_height: f32,
    font_metrics: Option<FontMetricSample>,
) {
    for (visual_line_index, run) in buffer.layout_runs().enumerate() {
        let _ = font_alignment;
        let line_y = if vertical_text {
            vertical_line_offsets
                .get(visual_line_index)
                .map(|offset| offset + run.line_y - run.line_top)
                .unwrap_or(run.line_y)
        } else {
            run.line_y
        };
        let baseline_y = ((origin_y + line_y) * device_scale).round() as i32;
        #[cfg(debug_assertions)]
        let mut ink_min_y = f32::INFINITY;
        #[cfg(debug_assertions)]
        let mut ink_max_y = f32::NEG_INFINITY;
        let rotate_run = vertical_text && should_rotate_vertical_run(run.text);
        let transform_run = rotate_run;
        let mut rotated_pixels = Vec::new();
        for (glyph_index, glyph) in run.glyphs.iter().enumerate() {
            let cluster = &run.text[glyph.start..glyph.end];
            let line_offset = if visual_line_index == 0 {
                first_line_offset
            } else {
                subsequent_line_offset
            };
            let hanging_offset = if hanging_punctuation {
                let starts_with_hanging = glyph_index == 0
                    && cluster
                        .chars()
                        .next()
                        .is_some_and(is_hanging_opening_punctuation);
                let ends_with_hanging = glyph_index + 1 == run.glyphs.len()
                    && cluster
                        .chars()
                        .last()
                        .is_some_and(is_hanging_closing_punctuation);
                if starts_with_hanging {
                    -glyph.w * 0.5
                } else if ends_with_hanging {
                    // cosmic-text already advances the closing punctuation after
                    // the preceding glyph. Moving it right again creates a visible
                    // gap before a final CJK full stop. Keep the glyph adjacent;
                    // the line layout still retains its original advance.
                    0.0
                } else {
                    0.0
                }
            } else {
                0.0
            };
            let physical = glyph.physical(
                (
                    (origin_x + line_offset + hanging_offset) * device_scale,
                    0.0,
                ),
                device_scale,
            );
            let color = glyph.color_opt.unwrap_or_else(|| Color::rgb(0, 0, 0));
            if transform_run {
                swash_cache.with_pixels(
                    font_system,
                    physical.cache_key,
                    color,
                    |offset_x, offset_y, pixel_color| {
                        rotated_pixels.push((
                            physical.x + offset_x,
                            baseline_y + physical.y + offset_y,
                            pixel_color,
                        ));
                    },
                );
            } else {
                swash_cache.with_pixels(
                    font_system,
                    physical.cache_key,
                    color,
                    |offset_x, offset_y, pixel_color| {
                        let x = physical.x + offset_x;
                        let y = baseline_y + physical.y + offset_y;
                        #[cfg(debug_assertions)]
                        {
                            ink_min_y = ink_min_y.min((physical.y + offset_y) as f32);
                            ink_max_y = ink_max_y.max((physical.y + offset_y) as f32);
                        }
                        if x < 0 || y < 0 || x >= bitmap_width as i32 || y >= bitmap_height as i32 {
                            return;
                        }
                        let index = ((y as u32 * bitmap_width + x as u32) * 4) as usize;
                        blend_pixel(pixels, index, pixel_color);
                    },
                );
            }
        }
        if transform_run && !rotated_pixels.is_empty() {
            let min_x = rotated_pixels.iter().map(|(x, _, _)| *x).min().unwrap_or(0);
            let max_x = rotated_pixels.iter().map(|(x, _, _)| *x).max().unwrap_or(0);
            let min_y = rotated_pixels.iter().map(|(_, y, _)| *y).min().unwrap_or(0);
            let max_y = rotated_pixels.iter().map(|(_, y, _)| *y).max().unwrap_or(0);
            let center_x = (min_x + max_x) / 2;
            let center_y = (min_y + max_y) / 2;
            let target_center_x = if vertical_text && column_width > 0.0 {
                ((origin_x + column_width / 2.0) * device_scale).round() as i32
            } else {
                center_x
            };
            let target_center_y = if vertical_text {
                let slot_start = vertical_line_offsets
                    .get(visual_line_index)
                    .copied()
                    .unwrap_or(0.0);
                let slot_end = vertical_line_offsets
                    .get(visual_line_index + 1)
                    .copied()
                    .unwrap_or(vertical_column_height);
                ((origin_y + (slot_start + slot_end) / 2.0) * device_scale).round() as i32
            } else {
                center_y
            };
            for (x, y, pixel_color) in rotated_pixels {
                let dx = x - center_x;
                let dy = y - center_y;
                // Rotate the shaped Latin token as a unit. Its advance is
                // measured from the real font, so no fixed scale is applied
                // to the glyph pixels here.
                let (mut rotated_x, rotated_y) = (center_x - dy, center_y + dx);
                rotated_x += target_center_x - center_x;
                let rotated_y = rotated_y + target_center_y - center_y;
                if rotated_x < 0
                    || rotated_y < 0
                    || rotated_x >= bitmap_width as i32
                    || rotated_y >= bitmap_height as i32
                {
                    continue;
                }
                let index = ((rotated_y as u32 * bitmap_width + rotated_x as u32) * 4) as usize;
                blend_pixel(pixels, index, pixel_color);
            }
        }
        #[cfg(debug_assertions)]
        {
            let metrics = font_metrics
                .map(|value| {
                    format!(
                        "ascent={:.2} descent={:.2} leading={:.2}",
                        value.ascent, value.descent, value.leading
                    )
                })
                .unwrap_or_else(|| "ascent=? descent=? leading=?".to_string());
            let ink_height = if ink_min_y.is_finite() {
                (ink_max_y - ink_min_y + 1.0) / device_scale
            } else {
                0.0
            };
            let preview: String = run.text.chars().take(24).collect();
            web_sys::console::log_1(&wasm_bindgen::JsValue::from_str(&format!(
                "[TextLayout] text={:?} line={} lineWidth={:.2} glyphs={} fontSize={:.2} {} metricLineHeight={:.2} requestedLineHeight={:.2} layoutLineHeight={:.2} lineY={:.2} lineTop={:.2} inkHeight={:.2}",
                preview,
                visual_line_index,
                run.line_w,
                run.glyphs.len(),
                font_size,
                metrics,
                metric_line_height,
                requested_line_height,
                run.line_height,
                run.line_y,
                run.line_top,
                ink_height
            )));
        }
    }
}

fn should_rotate_vertical_run(text: &str) -> bool {
    text.chars().any(is_vertical_rotated_char)
}

fn is_hanging_opening_punctuation(ch: char) -> bool {
    matches!(
        ch,
        '"' | '\''
            | '('
            | '['
            | '{'
            | '<'
            | '“'
            | '‘'
            | '（'
            | '【'
            | '《'
            | '「'
            | '『'
            | '〈'
            | '〔'
            | '［'
            | '｛'
    )
}

fn is_hanging_closing_punctuation(ch: char) -> bool {
    matches!(
        ch,
        ',' | '.'
            | ':'
            | ';'
            | '!'
            | '?'
            | ')'
            | ']'
            | '}'
            | '>'
            | '，'
            | '。'
            | '、'
            | '：'
            | '；'
            | '！'
            | '？'
            | '）'
            | '】'
            | '》'
            | '」'
            | '』'
            | '〉'
            | '〕'
            | '］'
            | '｝'
    )
}
