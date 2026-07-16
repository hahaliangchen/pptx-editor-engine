use crate::ast::ImageElement;
use wasm_bindgen::{prelude::*, JsCast};
use web_sys::{CanvasRenderingContext2d, HtmlImageElement};

pub fn render_image(
    ctx: &CanvasRenderingContext2d,
    image: &ImageElement,
    images_obj: &JsValue,
) -> Result<(), JsValue> {
    ctx.save();
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
                if let Some(crop) = &image.crop {
                    let source_width = html_image.natural_width() as f64;
                    let source_height = html_image.natural_height() as f64;
                    let sx = source_width * crop.left.clamp(0.0, 1.0) as f64;
                    let sy = source_height * crop.top.clamp(0.0, 1.0) as f64;
                    let sw = source_width * (1.0 - crop.left - crop.right).clamp(0.0, 1.0) as f64;
                    let sh = source_height * (1.0 - crop.top - crop.bottom).clamp(0.0, 1.0) as f64;
                    if sw > 0.0 && sh > 0.0 {
                        let _ = ctx.draw_image_with_html_image_element_and_sw_and_sh_and_dx_and_dy_and_dw_and_dh(
                            &html_image,
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
                        &html_image,
                        image.rect.x as f64,
                        image.rect.y as f64,
                        image.rect.w as f64,
                        image.rect.h as f64,
                    );
                }
            }
        }
    }
    ctx.restore();
    Ok(())
}
