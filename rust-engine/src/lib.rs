mod ast;

use ast::{
    Element, FillStyle, ParagraphStyle, Rect, ShapeElement, Slide, TextBodyProperties, TextElement,
    TextParagraph, TextRun, TextStyle,
};
use cosmic_text::fontdb::{self, Source};
use cosmic_text::{
    Align, Attrs, AttrsList, Buffer, BufferLine, CacheKeyFlags, Color, Family, FontSystem, Hinting,
    LineEnding, Metrics, Shaping, Style, SwashCache, Weight, Wrap,
};
use std::sync::Arc;
use wasm_bindgen::JsCast;
use wasm_bindgen::{prelude::*, Clamped};
use web_sys::{CanvasRenderingContext2d, HtmlCanvasElement, ImageData};

#[wasm_bindgen]
pub struct RustPptRenderer {
    ctx: CanvasRenderingContext2d,
    font_system: Option<FontSystem>,
    swash_cache: SwashCache,
    registered_fonts: Vec<Arc<Vec<u8>>>,
}

struct CosmicParagraph {
    buffer: Buffer,
    bullet_buffer: Option<Buffer>,
    x: f32,
    first_line_offset: f32,
    bullet_x: f32,
    top: f32,
    hanging_punctuation: bool,
    font_alignment: String,
    font_size: f32,
    metric_line_height: f32,
    line_height: f32,
    font_metrics: Option<FontMetricSample>,
}

#[derive(Clone, Copy)]
struct FontMetricSample {
    ascent: f32,
    descent: f32,
    leading: f32,
    line_height: f32,
}

// The mac-like profile uses grayscale, unhinted glyphs and a small oversampling
// factor before compositing the text bitmap onto the presentation canvas.
const MAC_TEXT_OVERSAMPLE: f32 = 2.0;
// WPS leaves less internal leading than the raw cosmic-text alignment box for
// the Windows fonts used by this deck. This is renderer compatibility behavior,
// not an OOXML spacing value.
const WPS_FONT_METRIC_SCALE: f32 = 0.92;

impl RustPptRenderer {
    fn color_with_opacity(color: &str, opacity: f32) -> String {
        if color.len() == 7 && color.starts_with('#') {
            if let (Ok(r), Ok(g), Ok(b)) = (
                u8::from_str_radix(&color[1..3], 16),
                u8::from_str_radix(&color[3..5], 16),
                u8::from_str_radix(&color[5..7], 16),
            ) {
                return format!("rgba({},{},{},{})", r, g, b, opacity.clamp(0.0, 1.0));
            }
        }
        color.to_string()
    }

    fn set_shape_paint(&self, fill: &FillStyle, rect: &Rect, stroke: bool) -> bool {
        match fill {
            FillStyle::None => false,
            FillStyle::Solid { color } => {
                if stroke {
                    self.ctx.set_stroke_style_str(color);
                } else {
                    self.ctx.set_fill_style_str(color);
                }
                true
            }
            FillStyle::Gradient {
                kind, stops, angle, ..
            } => {
                if stops.is_empty() {
                    return false;
                }
                let x = rect.x as f64;
                let y = rect.y as f64;
                let w = rect.w as f64;
                let h = rect.h as f64;
                let gradient = if kind == "radial" {
                    let radius = w.max(h) / 2.0;
                    self.ctx
                        .create_radial_gradient(
                            x + w / 2.0,
                            y + h / 2.0,
                            0.0,
                            x + w / 2.0,
                            y + h / 2.0,
                            radius,
                        )
                        .ok()
                } else {
                    let radians = angle.unwrap_or(0.0) as f64 * std::f64::consts::PI / 180.0;
                    let dx = radians.cos() * w / 2.0;
                    let dy = radians.sin() * h / 2.0;
                    Some(self.ctx.create_linear_gradient(
                        x + w / 2.0 - dx,
                        y + h / 2.0 - dy,
                        x + w / 2.0 + dx,
                        y + h / 2.0 + dy,
                    ))
                };
                let Some(gradient) = gradient else {
                    return false;
                };
                for stop in stops {
                    let _ = gradient.add_color_stop(stop.position, &stop.color);
                }
                if stroke {
                    self.ctx.set_stroke_style_canvas_gradient(&gradient);
                } else {
                    self.ctx.set_fill_style_canvas_gradient(&gradient);
                }
                true
            }
        }
    }

    fn begin_shape_path(&self, shp: &ShapeElement) {
        let x = shp.rect.x as f64;
        let y = shp.rect.y as f64;
        let w = shp.rect.w as f64;
        let h = shp.rect.h as f64;
        self.ctx.begin_path();
        match shp.shape_type.as_str() {
            "roundRect" => {
                let radius = (w.min(h) * 0.13).max(1.0);
                self.ctx.move_to(x + radius, y);
                self.ctx.line_to(x + w - radius, y);
                self.ctx.quadratic_curve_to(x + w, y, x + w, y + radius);
                self.ctx.line_to(x + w, y + h - radius);
                self.ctx
                    .quadratic_curve_to(x + w, y + h, x + w - radius, y + h);
                self.ctx.line_to(x + radius, y + h);
                self.ctx.quadratic_curve_to(x, y + h, x, y + h - radius);
                self.ctx.line_to(x, y + radius);
                self.ctx.quadratic_curve_to(x, y, x + radius, y);
            }
            "ellipse" => {
                let _ = self.ctx.ellipse(
                    x + w / 2.0,
                    y + h / 2.0,
                    w / 2.0,
                    h / 2.0,
                    0.0,
                    0.0,
                    2.0 * std::f64::consts::PI,
                );
            }
            "triangle" => {
                self.ctx.move_to(x + w / 2.0, y);
                self.ctx.line_to(x + w, y + h);
                self.ctx.line_to(x, y + h);
            }
            _ => self.ctx.rect(x, y, w, h),
        }
        self.ctx.close_path();
    }

    fn apply_shape_effects(&self, shp: &ShapeElement) {
        let Some(effects) = shp
            .computed_style
            .as_ref()
            .and_then(|style| style.effects.as_ref())
        else {
            return;
        };
        if let Some(shadow) = &effects.outer_shadow {
            let radians = shadow.direction as f64 * std::f64::consts::PI / 180.0;
            self.ctx
                .set_shadow_color(&Self::color_with_opacity(&shadow.color, shadow.opacity));
            self.ctx.set_shadow_blur(shadow.blur as f64);
            self.ctx
                .set_shadow_offset_x(radians.cos() * shadow.distance as f64);
            self.ctx
                .set_shadow_offset_y(radians.sin() * shadow.distance as f64);
        } else if let Some(glow) = &effects.glow {
            self.ctx
                .set_shadow_color(&Self::color_with_opacity(&glow.color, glow.opacity));
            self.ctx.set_shadow_blur(glow.radius as f64);
            self.ctx.set_shadow_offset_x(0.0);
            self.ctx.set_shadow_offset_y(0.0);
        }
    }

    fn clear_shape_effects(&self) {
        self.ctx.set_shadow_color("rgba(0,0,0,0)");
        self.ctx.set_shadow_blur(0.0);
        self.ctx.set_shadow_offset_x(0.0);
        self.ctx.set_shadow_offset_y(0.0);
    }

    fn apply_shadow_transform(&self, shp: &ShapeElement) {
        let Some(shadow) = shp
            .computed_style
            .as_ref()
            .and_then(|style| style.effects.as_ref())
            .and_then(|effects| effects.outer_shadow.as_ref())
        else {
            return;
        };
        let scale_x = shadow.scale_x.max(0.01) as f64;
        let scale_y = shadow.scale_y.max(0.01) as f64;
        if (scale_x - 1.0).abs() < f64::EPSILON && (scale_y - 1.0).abs() < f64::EPSILON {
            return;
        }
        let center_x = (shp.rect.x + shp.rect.w / 2.0) as f64;
        let center_y = (shp.rect.y + shp.rect.h / 2.0) as f64;
        let _ = self.ctx.translate(center_x, center_y);
        let _ = self.ctx.scale(scale_x, scale_y);
        let _ = self.ctx.translate(-center_x, -center_y);
    }

    fn paint_shape(&self, shp: &ShapeElement) {
        self.begin_shape_path(shp);
        if let Some(style) = &shp.computed_style {
            if self.set_shape_paint(&style.fill, &shp.rect, false) {
                self.ctx.fill();
            }
            if let Some(line) = &style.line {
                if self.set_shape_paint(&line.fill, &shp.rect, true) {
                    self.ctx.set_line_width(line.width as f64);
                    if let Some(cap) = &line.cap {
                        self.ctx.set_line_cap(cap);
                    }
                    if let Some(join) = &line.join {
                        self.ctx.set_line_join(join);
                    }
                    self.ctx.stroke();
                }
            }
        } else {
            if shp.fill != "transparent" {
                self.ctx.set_fill_style_str(&shp.fill);
                self.ctx.fill();
            }
            if let Some(border) = &shp.border {
                self.ctx.set_stroke_style_str(&border.color);
                self.ctx.set_line_width(border.width as f64);
                self.ctx.stroke();
            }
        }
    }
    fn is_east_asian_text(text: &str) -> bool {
        text.chars().any(|ch| {
            matches!(
                ch as u32,
                0x2E80..=0x2FFF
                    | 0x3040..=0x30FF
                    | 0x3100..=0x312F
                    | 0x3130..=0x318F
                    | 0x31A0..=0x31BF
                    | 0x31F0..=0x31FF
                    | 0x3400..=0x4DBF
                    | 0x4E00..=0x9FFF
                    | 0xAC00..=0xD7AF
                    | 0xF900..=0xFAFF
            )
        })
    }

    fn parse_text_color(value: &str) -> Color {
        if value.len() == 7 && value.starts_with('#') {
            if let (Ok(r), Ok(g), Ok(b)) = (
                u8::from_str_radix(&value[1..3], 16),
                u8::from_str_radix(&value[3..5], 16),
                u8::from_str_radix(&value[5..7], 16),
            ) {
                return Color::rgb(r, g, b);
            }
        }
        if let Some(channels) = value
            .strip_prefix("rgba(")
            .and_then(|inner| inner.strip_suffix(')'))
        {
            let parts: Vec<&str> = channels.split(',').map(str::trim).collect();
            if parts.len() == 4 {
                if let (Ok(r), Ok(g), Ok(b), Ok(alpha)) = (
                    parts[0].parse::<u8>(),
                    parts[1].parse::<u8>(),
                    parts[2].parse::<u8>(),
                    parts[3].parse::<f32>(),
                ) {
                    return Color::rgba(r, g, b, (alpha.clamp(0.0, 1.0) * 255.0) as u8);
                }
            }
        }
        Color::rgb(0, 0, 0)
    }

    fn cosmic_attrs<'a>(
        style: &'a TextStyle,
        text: &str,
        scale: f32,
        line_height: f32,
    ) -> Attrs<'a> {
        let requested_family =
            if Self::is_east_asian_text(text) && !style.east_asian_font_family.is_empty() {
                &style.east_asian_font_family
            } else {
                &style.font_family
            };
        let family = match requested_family.to_ascii_lowercase().as_str() {
            "serif" => Family::Serif,
            "sans-serif" | "sans serif" => Family::SansSerif,
            "cursive" => Family::Cursive,
            "fantasy" => Family::Fantasy,
            "monospace" => Family::Monospace,
            _ => Family::Name(requested_family),
        };
        let mut attrs = Attrs::new()
            .family(family)
            .color(Self::parse_text_color(&style.color))
            .cache_key_flags(CacheKeyFlags::DISABLE_HINTING)
            .metrics(Metrics::new(
                (style.font_size * scale).max(1.0),
                line_height.max(style.font_size * scale).max(1.0),
            ));
        if style.bold {
            attrs = attrs.weight(Weight::BOLD);
        }
        if style.italic {
            attrs = attrs.style(Style::Italic);
        }
        if style.letter_spacing != 0.0 && style.font_size > 0.0 {
            attrs = attrs.letter_spacing(style.letter_spacing / style.font_size);
        }
        attrs
    }

    fn font_metric_line_height(
        font_system: &mut FontSystem,
        style: &TextStyle,
        text: &str,
        scale: f32,
    ) -> Option<FontMetricSample> {
        let font_size = (style.font_size * scale).max(1.0);
        let attrs = Self::cosmic_attrs(style, text, scale, font_size);
        let family = match attrs.family {
            Family::Name(name) => fontdb::Family::Name(name),
            Family::Serif => fontdb::Family::Serif,
            Family::SansSerif => fontdb::Family::SansSerif,
            Family::Cursive => fontdb::Family::Cursive,
            Family::Fantasy => fontdb::Family::Fantasy,
            Family::Monospace => fontdb::Family::Monospace,
        };
        let families = [family];
        let id = font_system.db().query(&fontdb::Query {
            families: &families,
            weight: attrs.weight,
            stretch: attrs.stretch,
            style: attrs.style,
        })?;
        let font = font_system.get_font(id, attrs.weight)?;
        let metrics = font.metrics();
        if metrics.units_per_em == 0 {
            return None;
        }
        let px_scale = font_size / metrics.units_per_em as f32;
        let ascent = metrics.ascent * px_scale;
        let descent = metrics.descent * px_scale;
        let leading = metrics.leading * px_scale;
        let line_height = ascent - descent + leading;
        Some(FontMetricSample {
            ascent,
            descent,
            leading,
            line_height: (line_height * WPS_FONT_METRIC_SCALE).max(font_size),
        })
    }

    fn normalize_text_element(txt: &TextElement) -> TextElement {
        if !txt.paragraphs.is_empty() {
            return txt.clone();
        }
        let paragraphs = txt
            .content
            .split('\n')
            .map(|content| TextParagraph {
                runs: vec![TextRun {
                    content: content.to_string(),
                    style: txt.style.clone(),
                }],
                style: ParagraphStyle {
                    align: txt.style.align.clone(),
                    level: 0,
                    margin_left: 0.0,
                    indent: 0.0,
                    east_asian_line_break: false,
                    hanging_punctuation: false,
                    font_alignment: "auto".to_string(),
                    line_spacing: None,
                    space_before: 0.0,
                    space_after: 0.0,
                },
                bullet: None,
            })
            .collect();
        TextElement {
            id: txt.id.clone(),
            rect: txt.rect.clone(),
            content: txt.content.clone(),
            style: txt.style.clone(),
            paragraphs,
            body: txt.body.clone().or_else(|| {
                Some(TextBodyProperties {
                    margin_left: 8.0,
                    margin_right: 8.0,
                    margin_top: 4.0,
                    margin_bottom: 4.0,
                    vertical_anchor: "top".to_string(),
                    auto_fit: "none".to_string(),
                    font_scale: 1.0,
                })
            }),
        }
    }

    fn build_cosmic_paragraphs(
        font_system: &mut FontSystem,
        txt: &TextElement,
        scale: f32,
    ) -> (Vec<CosmicParagraph>, f32) {
        let (margin_left, margin_right, margin_top, margin_bottom) = txt
            .body
            .as_ref()
            .map(|body| {
                (
                    body.margin_left,
                    body.margin_right,
                    body.margin_top,
                    body.margin_bottom,
                )
            })
            .unwrap_or((8.0, 8.0, 4.0, 4.0));
        let inner_width = (txt.rect.w - margin_left - margin_right).max(1.0);
        let mut cursor_y = margin_top;
        let mut layouts = Vec::new();

        for paragraph in &txt.paragraphs {
            cursor_y += paragraph.style.space_before * scale;
            let first_style = paragraph
                .runs
                .first()
                .map(|run| &run.style)
                .unwrap_or(&txt.style);
            let font_size = paragraph
                .runs
                .iter()
                .map(|run| run.style.font_size * scale)
                .fold(first_style.font_size * scale, f32::max)
                .max(1.0);
            // A PPT percentage line spacing is based on the font's alignment box,
            // not only on the em-sized glyph. This preserves ascent/descent/leading
            // from the backend font and is the difference visible in WPS comparisons.
            let font_metrics = paragraph
                .runs
                .iter()
                .filter_map(|run| {
                    Self::font_metric_line_height(font_system, &run.style, &run.content, scale)
                })
                .max_by(|left, right| {
                    left.line_height
                        .partial_cmp(&right.line_height)
                        .unwrap_or(std::cmp::Ordering::Equal)
                });
            let metric_line_height = font_metrics
                .as_ref()
                .map(|metrics| metrics.line_height)
                .unwrap_or(font_size);
            let line_height = paragraph
                .style
                .line_spacing
                .as_ref()
                .map(|spacing| {
                    if spacing.unit == "points" {
                        (spacing.value * scale).max(font_size)
                    } else {
                        (metric_line_height * spacing.value).max(font_size)
                    }
                })
                // PPT paragraphs without an explicit lnSpc use the font's line box.
                // Explicit spcPct/spcPts values are handled above.
                .unwrap_or(metric_line_height);

            let mut text = String::new();
            for run in &paragraph.runs {
                text.push_str(&run.content);
            }
            let default_attrs = Self::cosmic_attrs(first_style, &text, scale, line_height);
            let mut attrs_list = AttrsList::new(&default_attrs);
            let mut byte_offset = 0;
            for run in &paragraph.runs {
                let end = byte_offset + run.content.len();
                let run_attrs = Self::cosmic_attrs(&run.style, &run.content, scale, line_height);
                attrs_list.add_span(byte_offset..end, &run_attrs);
                byte_offset = end;
            }

            let paragraph_left = paragraph.style.margin_left.max(0.0);
            let text_indent = if paragraph.bullet.is_none() {
                paragraph.style.indent.max(0.0)
            } else {
                0.0
            };
            let x = margin_left + paragraph_left;
            let first_line_offset = if paragraph.bullet.is_none() {
                paragraph.style.indent
            } else {
                0.0
            };
            let available_width = (inner_width - paragraph_left - text_indent).max(1.0);
            let mut line = BufferLine::new(&text, LineEnding::None, attrs_list, Shaping::Advanced);
            line.set_align(Some(match paragraph.style.align.as_str() {
                "center" => Align::Center,
                "right" => Align::Right,
                _ => Align::Left,
            }));
            let mut buffer = Buffer::new_empty(Metrics::new(font_size, line_height));
            buffer.lines.push(line);
            buffer.set_hinting(Hinting::Disabled);
            buffer.set_size(Some(available_width), Some(100_000.0));
            // PowerPoint's eaLnBrk allows CJK glyph boundaries as legal break points.
            // Word-only wrapping leaves long Chinese runs on one line or breaks them at
            // different places from Office, which changes the total text-box height.
            buffer.set_wrap(
                if paragraph.style.east_asian_line_break || Self::is_east_asian_text(&text) {
                    Wrap::WordOrGlyph
                } else {
                    Wrap::Word
                },
            );
            buffer.shape_until_scroll(font_system, false);

            let paragraph_height = buffer
                .layout_runs()
                .map(|run| run.line_top + run.line_height)
                .fold(line_height, f32::max);

            let bullet_buffer = paragraph.bullet.as_ref().map(|bullet| {
                let mut bullet_style = first_style.clone();
                bullet_style.color = bullet.color.clone();
                let bullet_attrs =
                    Self::cosmic_attrs(&bullet_style, &bullet.char, scale, line_height);
                let attrs = AttrsList::new(&bullet_attrs);
                let bullet_line =
                    BufferLine::new(&bullet.char, LineEnding::None, attrs, Shaping::Advanced);
                let mut bullet_buffer = Buffer::new_empty(Metrics::new(font_size, line_height));
                bullet_buffer.lines.push(bullet_line);
                bullet_buffer.set_hinting(Hinting::Disabled);
                bullet_buffer.set_size(Some(font_size * 4.0), Some(line_height * 2.0));
                bullet_buffer.set_wrap(Wrap::None);
                bullet_buffer.shape_until_scroll(font_system, false);
                bullet_buffer
            });

            layouts.push(CosmicParagraph {
                buffer,
                bullet_buffer,
                x,
                first_line_offset,
                bullet_x: margin_left + paragraph_left + paragraph.style.indent,
                top: cursor_y,
                hanging_punctuation: paragraph.style.hanging_punctuation,
                font_alignment: paragraph.style.font_alignment.clone(),
                font_size,
                metric_line_height,
                line_height,
                font_metrics,
            });
            cursor_y += paragraph_height + paragraph.style.space_after * scale;
        }

        (layouts, cursor_y + margin_bottom)
    }

    fn blend_pixel(target: &mut [u8], index: usize, color: Color) {
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

    #[allow(clippy::too_many_arguments)]
    fn rasterize_buffer(
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
        hanging_punctuation: bool,
        font_alignment: &str,
        font_size: f32,
        metric_line_height: f32,
        requested_line_height: f32,
        font_metrics: Option<FontMetricSample>,
    ) {
        for (visual_line_index, run) in buffer.layout_runs().enumerate() {
            // cosmic-text's line_y is computed from the actual font ascent/descent and
            // centers that glyph box inside the requested line height. The current sample
            // uses fontAlgn="auto"; keep the metric-derived baseline instead of inventing
            // offsets for values whose glyph metrics are not exposed by LayoutRun.
            let _ = font_alignment;
            let baseline_y = ((origin_y + run.line_y) * device_scale).round() as i32;
            #[cfg(debug_assertions)]
            let mut ink_min_y = f32::INFINITY;
            #[cfg(debug_assertions)]
            let mut ink_max_y = f32::NEG_INFINITY;
            for (glyph_index, glyph) in run.glyphs.iter().enumerate() {
                let cluster = &run.text[glyph.start..glyph.end];
                let line_offset = if visual_line_index == 0 {
                    first_line_offset
                } else {
                    0.0
                };
                let hanging_offset = if hanging_punctuation {
                    let starts_with_hanging = glyph_index == 0
                        && cluster
                            .chars()
                            .next()
                            .is_some_and(Self::is_hanging_opening_punctuation);
                    let ends_with_hanging = glyph_index + 1 == run.glyphs.len()
                        && cluster
                            .chars()
                            .last()
                            .is_some_and(Self::is_hanging_closing_punctuation);
                    if starts_with_hanging {
                        -glyph.w * 0.5
                    } else if ends_with_hanging {
                        glyph.w * 0.5
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
                        Self::blend_pixel(pixels, index, pixel_color);
                    },
                );
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
                web_sys::console::log_1(&JsValue::from_str(&format!(
                    "[TextLayout] text={:?} line={} fontSize={:.2} {} metricLineHeight={:.2} requestedLineHeight={:.2} layoutLineHeight={:.2} lineY={:.2} lineTop={:.2} inkHeight={:.2}",
                    preview,
                    visual_line_index,
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

    fn render_rich_text(&mut self, txt: &TextElement) -> Result<(), JsValue> {
        let body = txt.body.as_ref();
        let available_height = txt.rect.h.max(1.0);
        let mut scale = body
            .map(|value| value.font_scale)
            .unwrap_or(1.0)
            .clamp(0.2, 1.0);
        let transform = self.ctx.get_transform()?;
        let device_scale = ((transform.a() * transform.a() + transform.b() * transform.b()).sqrt()
            as f32)
            .max(0.1);
        let raster_scale = device_scale * MAC_TEXT_OVERSAMPLE;
        let bitmap_width = (txt.rect.w.max(1.0) * raster_scale).ceil() as u32;
        let bitmap_height = (txt.rect.h.max(1.0) * raster_scale).ceil() as u32;

        let (font_system_opt, swash_cache) = (&mut self.font_system, &mut self.swash_cache);
        let font_system = font_system_opt.as_mut().ok_or_else(|| {
            JsValue::from_str("Rich text rendering requires backend fonts to be registered")
        })?;
        let (mut paragraphs, mut layout_height) =
            Self::build_cosmic_paragraphs(font_system, txt, scale);

        if body.map(|value| value.auto_fit.as_str()) == Some("shrink") {
            while layout_height > available_height && scale > 0.2 {
                scale = (scale - 0.05).max(0.2);
                (paragraphs, layout_height) =
                    Self::build_cosmic_paragraphs(font_system, txt, scale);
            }
        }

        let vertical_offset = match body.map(|value| value.vertical_anchor.as_str()) {
            Some("middle") => (available_height - layout_height).max(0.0) / 2.0,
            Some("bottom") => (available_height - layout_height).max(0.0),
            _ => 0.0,
        };

        let mut pixels = vec![0_u8; bitmap_width as usize * bitmap_height as usize * 4];
        for paragraph in &mut paragraphs {
            let top = paragraph.top + vertical_offset;
            Self::rasterize_buffer(
                font_system,
                swash_cache,
                &paragraph.buffer,
                paragraph.x,
                top,
                raster_scale,
                &mut pixels,
                bitmap_width,
                bitmap_height,
                paragraph.first_line_offset,
                paragraph.hanging_punctuation,
                &paragraph.font_alignment,
                paragraph.font_size,
                paragraph.metric_line_height,
                paragraph.line_height,
                paragraph.font_metrics,
            );
            if let Some(bullet_buffer) = &paragraph.bullet_buffer {
                Self::rasterize_buffer(
                    font_system,
                    swash_cache,
                    bullet_buffer,
                    paragraph.bullet_x,
                    top,
                    raster_scale,
                    &mut pixels,
                    bitmap_width,
                    bitmap_height,
                    0.0,
                    false,
                    &paragraph.font_alignment,
                    paragraph.font_size,
                    paragraph.metric_line_height,
                    paragraph.line_height,
                    paragraph.font_metrics,
                );
            }
        }

        let document = web_sys::window()
            .and_then(|window| window.document())
            .ok_or_else(|| {
                JsValue::from_str("Document is unavailable for text bitmap compositing")
            })?;
        let canvas: HtmlCanvasElement = document.create_element("canvas")?.dyn_into()?;
        canvas.set_width(bitmap_width);
        canvas.set_height(bitmap_height);
        let bitmap_ctx: CanvasRenderingContext2d = canvas
            .get_context("2d")?
            .ok_or_else(|| JsValue::from_str("Could not create text bitmap context"))?
            .dyn_into()?;
        let image_data = ImageData::new_with_u8_clamped_array_and_sh(
            Clamped(&pixels),
            bitmap_width,
            bitmap_height,
        )?;
        bitmap_ctx.put_image_data(&image_data, 0.0, 0.0)?;
        self.ctx.draw_image_with_html_canvas_element_and_dw_and_dh(
            &canvas,
            txt.rect.x as f64,
            txt.rect.y as f64,
            txt.rect.w as f64,
            txt.rect.h as f64,
        )?;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::RustPptRenderer;

    #[test]
    fn parses_text_colors_for_rasterization() {
        let solid = RustPptRenderer::parse_text_color("#c00000");
        assert_eq!(
            (solid.r(), solid.g(), solid.b(), solid.a()),
            (192, 0, 0, 255)
        );

        let translucent = RustPptRenderer::parse_text_color("rgba(1,2,3,0.5)");
        assert_eq!(
            (translucent.r(), translucent.g(), translucent.b()),
            (1, 2, 3)
        );
        assert!((126..=128).contains(&translucent.a()));
    }

    #[test]
    fn blends_rasterized_glyph_pixels() {
        let mut target = [0_u8; 4];
        RustPptRenderer::blend_pixel(&mut target, 0, cosmic_text::Color::rgba(10, 20, 30, 128));
        assert_eq!((target[0], target[1], target[2]), (10, 20, 30));
        assert!((127..=129).contains(&target[3]));
    }
}
#[wasm_bindgen]
impl RustPptRenderer {
    #[wasm_bindgen(constructor)]
    pub fn new(ctx: CanvasRenderingContext2d) -> Self {
        std::panic::set_hook(Box::new(console_error_panic_hook::hook));
        Self {
            ctx,
            font_system: None,
            swash_cache: SwashCache::new(),
            registered_fonts: Vec::new(),
        }
    }

    #[wasm_bindgen]
    pub fn register_font(&mut self, font_bytes: &[u8]) {
        let arc_bytes = Arc::new(font_bytes.to_vec());
        self.registered_fonts.push(arc_bytes);

        let sources = self
            .registered_fonts
            .iter()
            .map(|f| Source::Binary(f.clone()))
            .collect::<Vec<_>>();

        // Re-create the font system with the registered fonts
        let mut fs = FontSystem::new_with_fonts(sources.into_iter());
        let fallback_family = fs
            .db()
            .faces()
            .next()
            .and_then(|face| face.families.first().map(|family| family.0.clone()));
        if let Some(family) = fallback_family {
            fs.db_mut().set_sans_serif_family(&family);
            fs.db_mut().set_serif_family(&family);
            fs.db_mut().set_monospace_family(&family);
        }

        // Log loaded font faces to the browser console for debugging
        for face in fs.db().faces() {
            let families_str = face
                .families
                .iter()
                .map(|f| &f.0)
                .cloned()
                .collect::<Vec<_>>()
                .join(", ");
            let msg = format!(
                "Loaded Font Face: Families=[{}], Weight={:?}, Style={:?}",
                families_str, face.weight, face.style
            );
            let _ = web_sys::console::log_1(&JsValue::from_str(&msg));
        }

        self.font_system = Some(fs);
        self.swash_cache = SwashCache::new();
    }

    // Render single slide from JSON AST, with images_obj mapping URLs to HTMLImageElements
    #[wasm_bindgen]
    pub fn render_slide(&mut self, slide_json: &str, images_obj: &JsValue) -> Result<(), JsValue> {
        // Debug: Log all currently registered fonts in the WASM font database
        if let Some(fs) = &self.font_system {
            let count = fs.db().faces().count();
            let mut face_names = Vec::new();
            for face in fs.db().faces() {
                let fams = face
                    .families
                    .iter()
                    .map(|f| &f.0)
                    .cloned()
                    .collect::<Vec<_>>()
                    .join(",");
                face_names.push(format!("{} (weight: {:?})", fams, face.weight));
            }
            let msg = format!(
                "[WASM Debug] render_slide called. Active font_system exists. Total faces in database: {}. Faces: {:?}",
                count, face_names
            );
            let _ = web_sys::console::log_1(&JsValue::from_str(&msg));
        } else {
            let _ = web_sys::console::log_1(&JsValue::from_str(
                "[WASM Debug] render_slide called. self.font_system is NONE (no fonts loaded yet).",
            ));
        }

        let slide: Slide = serde_json::from_str(slide_json)
            .map_err(|e| JsValue::from_str(&format!("JSON Parse Error: {}", e)))?;

        // 1. Clear and fill background (reset transform to cover physical canvas size)
        self.ctx.save();
        let _ = self.ctx.set_transform(1.0, 0.0, 0.0, 1.0, 0.0, 0.0);
        if let Some(canvas) = self.ctx.canvas() {
            self.ctx.set_fill_style_str("#ffffff");
            self.ctx
                .fill_rect(0.0, 0.0, canvas.width() as f64, canvas.height() as f64);
        }
        self.ctx.restore();

        // 2. Render elements from bottom to top (Z-Index order)
        for element in &slide.elements {
            match element {
                Element::Shape(shp) => {
                    self.ctx.save();
                    let has_outer_shadow = shp
                        .computed_style
                        .as_ref()
                        .and_then(|style| style.effects.as_ref())
                        .and_then(|effects| effects.outer_shadow.as_ref())
                        .is_some();
                    let has_effect = shp
                        .computed_style
                        .as_ref()
                        .and_then(|style| style.effects.as_ref())
                        .map(|effects| effects.outer_shadow.is_some() || effects.glow.is_some())
                        .unwrap_or(false);

                    if has_effect {
                        // Office renders outerShdw/glow behind the shape. Draw that pass
                        // separately so the normal fill and outline do not dilute it.
                        self.ctx.save();
                        if has_outer_shadow {
                            self.apply_shadow_transform(shp);
                        }
                        self.apply_shape_effects(shp);
                        self.paint_shape(shp);
                        self.ctx.restore();
                        self.clear_shape_effects();
                    }

                    self.paint_shape(shp);
                    self.ctx.restore();
                }
                Element::Text(txt) => {
                    self.ctx.save();
                    self.ctx.begin_path();
                    self.ctx.rect(
                        txt.rect.x as f64,
                        txt.rect.y as f64,
                        txt.rect.w as f64,
                        txt.rect.h as f64,
                    );
                    self.ctx.clip();
                    let normalized = Self::normalize_text_element(txt);
                    self.render_rich_text(&normalized)?;
                    self.ctx.restore();
                }
                Element::Image(img) => {
                    self.ctx.save();
                    self.ctx.set_image_smoothing_enabled(true);
                    let _ = js_sys::Reflect::set(
                        self.ctx.as_ref(),
                        &JsValue::from_str("imageSmoothingQuality"),
                        &JsValue::from_str("high"),
                    );
                    if let Some(img_val) =
                        js_sys::Reflect::get(images_obj, &JsValue::from_str(&img.url)).ok()
                    {
                        if !img_val.is_undefined() && !img_val.is_null() {
                            if let Ok(html_img) = img_val.dyn_into::<web_sys::HtmlImageElement>() {
                                if let Some(crop) = &img.crop {
                                    let source_width = html_img.natural_width() as f64;
                                    let source_height = html_img.natural_height() as f64;
                                    let sx = source_width * crop.left.clamp(0.0, 1.0) as f64;
                                    let sy = source_height * crop.top.clamp(0.0, 1.0) as f64;
                                    let sw = source_width
                                        * (1.0 - crop.left - crop.right).clamp(0.0, 1.0) as f64;
                                    let sh = source_height
                                        * (1.0 - crop.top - crop.bottom).clamp(0.0, 1.0) as f64;
                                    if sw > 0.0 && sh > 0.0 {
                                        let _ = self.ctx.draw_image_with_html_image_element_and_sw_and_sh_and_dx_and_dy_and_dw_and_dh(
                                            &html_img,
                                            sx,
                                            sy,
                                            sw,
                                            sh,
                                            img.rect.x as f64,
                                            img.rect.y as f64,
                                            img.rect.w as f64,
                                            img.rect.h as f64,
                                        );
                                    }
                                } else {
                                    let _ =
                                        self.ctx.draw_image_with_html_image_element_and_dw_and_dh(
                                            &html_img,
                                            img.rect.x as f64,
                                            img.rect.y as f64,
                                            img.rect.w as f64,
                                            img.rect.h as f64,
                                        );
                                }
                            }
                        }
                    }
                    self.ctx.restore();
                }
            }
        }
        Ok(())
    }
}
