use crate::ast::ShapeElement;
use wasm_bindgen::{prelude::*, Clamped};
use web_sys::{CanvasRenderingContext2d, HtmlCanvasElement, ImageData};

pub fn clear_canvas_shadow(ctx: &CanvasRenderingContext2d) {
    ctx.set_shadow_color("rgba(0,0,0,0)");
    ctx.set_shadow_blur(0.0);
    ctx.set_shadow_offset_x(0.0);
    ctx.set_shadow_offset_y(0.0);
}

fn shape_contains(shp: &ShapeElement, x: f32, y: f32) -> bool {
    // The effect bitmap is sampled in page space, while the shape path is
    // defined in its unrotated local box. Undo xfrm before testing coverage.
    let center_x = shp.rect.x + shp.rect.w / 2.0;
    let center_y = shp.rect.y + shp.rect.h / 2.0;
    let dx = x - center_x;
    let dy = y - center_y;
    let radians = -(shp.rotation * std::f32::consts::PI / 180.0);
    let rotated_x = dx * radians.cos() - dy * radians.sin();
    let rotated_y = dx * radians.sin() + dy * radians.cos();
    let local_x = if shp.flip_h { -rotated_x } else { rotated_x } + shp.rect.w / 2.0;
    let local_y = if shp.flip_v { -rotated_y } else { rotated_y } + shp.rect.h / 2.0;
    let width = shp.rect.w;
    let height = shp.rect.h;
    if local_x < 0.0 || local_y < 0.0 || local_x > width || local_y > height {
        return false;
    }

    match shp.shape_type.as_str() {
        "ellipse" => {
            let dx = (local_x - width / 2.0) / (width / 2.0).max(0.001);
            let dy = (local_y - height / 2.0) / (height / 2.0).max(0.001);
            dx * dx + dy * dy <= 1.0
        }
        "triangle" => {
            if height <= 0.0 {
                return false;
            }
            let half_width = (local_y / height).clamp(0.0, 1.0) * width / 2.0;
            (local_x - width / 2.0).abs() <= half_width
        }
        "mathPlus" => {
            let arm = width.min(height) / 3.0;
            let left = (width - arm) / 2.0;
            let top = (height - arm) / 2.0;
            (local_x >= left && local_x <= left + arm) || (local_y >= top && local_y <= top + arm)
        }
        "upArrow" => {
            let head_height = shp.arrow_head_height.unwrap_or(0.5).clamp(0.001, 0.5);
            let shaft_width = shp.arrow_shaft_width.unwrap_or(0.5).clamp(0.001, 1.0);
            let shaft_left = (width - width * shaft_width) / 2.0;
            let shaft_right = (width + width * shaft_width) / 2.0;
            let head_bottom = width.min(height) * head_height;
            (local_y <= head_bottom
                && local_x >= width / 2.0 - (local_y / head_bottom.max(0.001)) * width / 2.0
                && local_x <= width / 2.0 + (local_y / head_bottom.max(0.001)) * width / 2.0)
                || (local_y >= head_bottom && local_x >= shaft_left && local_x <= shaft_right)
        }
        "roundRect" => {
            let radius = (width.min(height)
                * shp.corner_radius.unwrap_or(1.0 / 6.0).clamp(0.0, 0.5))
            .min(width.min(height) / 2.0);
            if (local_x >= radius && local_x <= width - radius)
                || (local_y >= radius && local_y <= height - radius)
            {
                return true;
            }
            let corner_x = if local_x < radius {
                radius
            } else {
                width - radius
            };
            let corner_y = if local_y < radius {
                radius
            } else {
                height - radius
            };
            let dx = local_x - corner_x;
            let dy = local_y - corner_y;
            dx * dx + dy * dy <= radius * radius
        }
        "line" => false,
        _ => true,
    }
}

fn gaussian_blur_alpha(alpha: &[u8], width: usize, height: usize, sigma: f32) -> Vec<u8> {
    if alpha.is_empty() || width == 0 || height == 0 || sigma <= 0.01 {
        return alpha.to_vec();
    }
    let radius = (sigma * 3.0).ceil().max(1.0) as i32;
    let mut kernel = Vec::with_capacity((radius * 2 + 1) as usize);
    let mut kernel_sum = 0.0;
    for offset in -radius..=radius {
        let distance = offset as f32;
        let weight = (-distance * distance / (2.0 * sigma * sigma)).exp();
        kernel.push(weight);
        kernel_sum += weight;
    }
    for weight in &mut kernel {
        *weight /= kernel_sum;
    }

    let mut horizontal = vec![0.0_f32; alpha.len()];
    for y in 0..height {
        for x in 0..width {
            let mut value = 0.0;
            for (index, offset) in (-radius..=radius).enumerate() {
                let sample_x = (x as i32 + offset).clamp(0, width as i32 - 1) as usize;
                value += alpha[y * width + sample_x] as f32 * kernel[index];
            }
            horizontal[y * width + x] = value;
        }
    }

    let mut blurred = vec![0_u8; alpha.len()];
    for y in 0..height {
        for x in 0..width {
            let mut value = 0.0;
            for (index, offset) in (-radius..=radius).enumerate() {
                let sample_y = (y as i32 + offset).clamp(0, height as i32 - 1) as usize;
                value += horizontal[sample_y * width + x] * kernel[index];
            }
            blurred[y * width + x] = value.clamp(0.0, 255.0).round() as u8;
        }
    }
    blurred
}

fn parse_effect_rgb(color: &str) -> (u8, u8, u8) {
    if color.len() == 7 && color.starts_with('#') {
        if let (Ok(r), Ok(g), Ok(b)) = (
            u8::from_str_radix(&color[1..3], 16),
            u8::from_str_radix(&color[3..5], 16),
            u8::from_str_radix(&color[5..7], 16),
        ) {
            return (r, g, b);
        }
    }
    (0, 0, 0)
}

pub fn render_custom_shape_effect(
    ctx: &CanvasRenderingContext2d,
    shp: &ShapeElement,
    device_scale: f32,
) -> Result<(), JsValue> {
    let Some(effects) = shp
        .computed_style
        .as_ref()
        .and_then(|style| style.effects.as_ref())
    else {
        return Ok(());
    };

    let (color, opacity, blur, scale_x, scale_y, distance, direction) =
        if let Some(shadow) = &effects.outer_shadow {
            (
                shadow.color.as_str(),
                shadow.opacity,
                shadow.blur,
                shadow.scale_x.max(0.01),
                shadow.scale_y.max(0.01),
                shadow.distance,
                shadow.direction,
            )
        } else if let Some(glow) = &effects.glow {
            (
                glow.color.as_str(),
                glow.opacity,
                glow.radius,
                1.0,
                1.0,
                0.0,
                0.0,
            )
        } else {
            return Ok(());
        };

    let device_scale = device_scale.max(0.1);
    let blur_pixels = (blur * device_scale).max(0.0);
    let sigma = (blur_pixels * 0.5).max(0.5);
    let radians = direction * std::f32::consts::PI / 180.0;
    let offset_x = distance * radians.cos();
    let offset_y = distance * radians.sin();
    let scale_spread_x = (scale_x - 1.0).abs() * shp.rect.w / 2.0;
    let scale_spread_y = (scale_y - 1.0).abs() * shp.rect.h / 2.0;
    let padding = (sigma * 3.0 / device_scale)
        .max(offset_x.abs())
        .max(offset_y.abs())
        .max(scale_spread_x)
        .max(scale_spread_y)
        + 2.0 / device_scale;
    let origin_x = shp.rect.x - padding;
    let origin_y = shp.rect.y - padding;
    let logical_width = shp.rect.w + padding * 2.0;
    let logical_height = shp.rect.h + padding * 2.0;
    let bitmap_width = (logical_width * device_scale).ceil().max(1.0) as usize;
    let bitmap_height = (logical_height * device_scale).ceil().max(1.0) as usize;
    let center_x = shp.rect.x + shp.rect.w / 2.0;
    let center_y = shp.rect.y + shp.rect.h / 2.0;

    let mut mask = vec![0_u8; bitmap_width * bitmap_height];
    for y in 0..bitmap_height {
        for x in 0..bitmap_width {
            let mut covered = 0_u8;
            for sample_y in [0.25_f32, 0.75_f32] {
                for sample_x in [0.25_f32, 0.75_f32] {
                    let target_x = origin_x + (x as f32 + sample_x) / device_scale;
                    let target_y = origin_y + (y as f32 + sample_y) / device_scale;
                    let source_x = (target_x - offset_x - center_x) / scale_x + center_x;
                    let source_y = (target_y - offset_y - center_y) / scale_y + center_y;
                    if shape_contains(shp, source_x, source_y) {
                        covered += 1;
                    }
                }
            }
            mask[y * bitmap_width + x] = (covered as u16 * 64).min(255) as u8;
        }
    }

    let blurred = gaussian_blur_alpha(&mask, bitmap_width, bitmap_height, sigma);
    let (red, green, blue) = parse_effect_rgb(color);
    let opacity = opacity.clamp(0.0, 1.0);
    let mut pixels = vec![0_u8; bitmap_width * bitmap_height * 4];
    for (index, alpha) in blurred.iter().enumerate() {
        let output_index = index * 4;
        pixels[output_index] = red;
        pixels[output_index + 1] = green;
        pixels[output_index + 2] = blue;
        pixels[output_index + 3] = (*alpha as f32 * opacity).round() as u8;
    }

    let document = web_sys::window()
        .and_then(|window| window.document())
        .ok_or_else(|| JsValue::from_str("Document is unavailable for shape effects"))?;
    let canvas: HtmlCanvasElement = document.create_element("canvas")?.dyn_into()?;
    canvas.set_width(bitmap_width as u32);
    canvas.set_height(bitmap_height as u32);
    let bitmap_ctx: CanvasRenderingContext2d = canvas
        .get_context("2d")?
        .ok_or_else(|| JsValue::from_str("Could not create shape effect canvas"))?
        .dyn_into()?;
    let image_data = ImageData::new_with_u8_clamped_array_and_sh(
        Clamped(&pixels),
        bitmap_width as u32,
        bitmap_height as u32,
    )?;
    bitmap_ctx.put_image_data(&image_data, 0.0, 0.0)?;
    ctx.draw_image_with_html_canvas_element_and_dw_and_dh(
        &canvas,
        origin_x as f64,
        origin_y as f64,
        logical_width as f64,
        logical_height as f64,
    )?;
    Ok(())
}
