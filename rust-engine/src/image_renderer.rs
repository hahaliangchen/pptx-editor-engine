use crate::ast::{ImageElement, ShadowStyle};
use wasm_bindgen::{prelude::*, JsCast};
use web_sys::{CanvasRenderingContext2d, HtmlImageElement};

fn shadow_rgba(shadow: &ShadowStyle) -> String {
    if shadow.color.starts_with("rgba(") || shadow.color.starts_with("rgb(") {
        return shadow.color.clone();
    }
    let hex = shadow.color.trim_start_matches('#');
    if hex.len() == 6 {
        return format!(
            "rgba({},{},{},{})",
            u8::from_str_radix(&hex[0..2], 16).unwrap_or(0),
            u8::from_str_radix(&hex[2..4], 16).unwrap_or(0),
            u8::from_str_radix(&hex[4..6], 16).unwrap_or(0),
            shadow.opacity.clamp(0.0, 1.0)
        );
    }
    "rgba(0,0,0,0)".to_string()
}

fn draw_image_source(
    ctx: &CanvasRenderingContext2d,
    image: &ImageElement,
    html_image: &HtmlImageElement,
) {
    if let Some(crop) = &image.crop {
        let source_width = html_image.natural_width() as f64;
        let source_height = html_image.natural_height() as f64;
        let sx = source_width * crop.left.clamp(0.0, 1.0) as f64;
        let sy = source_height * crop.top.clamp(0.0, 1.0) as f64;
        let sw = source_width * (1.0 - crop.left - crop.right).clamp(0.0, 1.0) as f64;
        let sh = source_height * (1.0 - crop.top - crop.bottom).clamp(0.0, 1.0) as f64;
        if sw > 0.0 && sh > 0.0 {
            let _ = ctx
                .draw_image_with_html_image_element_and_sw_and_sh_and_dx_and_dy_and_dw_and_dh(
                    html_image,
                    sx,
                    sy,
                    sw,
                    sh,
                    image.rect.x as f64,
                    image.rect.y as f64,
                    image.rect.w as f64,
                    image.rect.h as f64,
                );
        }
    } else {
        let _ = ctx.draw_image_with_html_image_element_and_dw_and_dh(
            html_image,
            image.rect.x as f64,
            image.rect.y as f64,
            image.rect.w as f64,
            image.rect.h as f64,
        );
    }
}

pub fn render_image(
    ctx: &CanvasRenderingContext2d,
    image: &ImageElement,
    images_obj: &JsValue,
) -> Result<(), JsValue> {
    ctx.save();
    let center_x = (image.rect.x + image.rect.w / 2.0) as f64;
    let center_y = (image.rect.y + image.rect.h / 2.0) as f64;
    let _ = ctx.translate(center_x, center_y);
    let _ = ctx.rotate((image.rotation as f64).to_radians());
    let _ = ctx.scale(
        if image.flip_h { -1.0 } else { 1.0 },
        if image.flip_v { -1.0 } else { 1.0 },
    );
    let _ = ctx.translate(-center_x, -center_y);
    ctx.set_image_smoothing_enabled(true);
    let _ = js_sys::Reflect::set(
        ctx.as_ref(),
        &JsValue::from_str("imageSmoothingQuality"),
        &JsValue::from_str("high"),
    );

    if let Some(image_value) = js_sys::Reflect::get(images_obj, &JsValue::from_str(&image.url)).ok()
    {
        if !image_value.is_undefined() && !image_value.is_null() {
            if let Ok(html_image) = image_value.dyn_into::<HtmlImageElement>() {
                if let Some(shadow) = image
                    .effects
                    .as_ref()
                    .and_then(|effects| effects.outer_shadow.as_ref())
                {
                    let radians = shadow.direction.to_radians();
                    ctx.set_shadow_color(&shadow_rgba(shadow));
                    ctx.set_shadow_blur(shadow.blur.max(0.0) as f64);
                    ctx.set_shadow_offset_x((shadow.distance * radians.cos()) as f64);
                    ctx.set_shadow_offset_y((shadow.distance * radians.sin()) as f64);
                    draw_image_source(&ctx, image, &html_image);
                    ctx.set_shadow_color("rgba(0,0,0,0)");
                    ctx.set_shadow_blur(0.0);
                    ctx.set_shadow_offset_x(0.0);
                    ctx.set_shadow_offset_y(0.0);
                }
                draw_image_source(&ctx, image, &html_image);
            }
        }
    }
    ctx.restore();
    Ok(())
}
