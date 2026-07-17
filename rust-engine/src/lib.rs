mod ast;
mod effects;
mod font_renderer;
mod image_renderer;
mod shape_renderer;
mod text_layout;

use ast::{Element, ReflectionStyle, Slide, TextElement};
use cosmic_text::fontdb::Source;
use cosmic_text::{FontSystem, SwashCache};
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

// The mac-like profile uses grayscale, unhinted glyphs and a small oversampling
// factor before compositing the text bitmap onto the presentation canvas.
const MAC_TEXT_OVERSAMPLE: f32 = 2.0;

impl RustPptRenderer {
    fn text_reflection(txt: &TextElement) -> Option<ReflectionStyle> {
        txt.paragraphs
            .iter()
            .flat_map(|paragraph| paragraph.runs.iter())
            .find_map(|run| run.style.reflection.clone())
            .or_else(|| txt.style.reflection.clone())
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

        let (font_system_opt, swash_cache) = (&mut self.font_system, &mut self.swash_cache);
        let font_system = font_system_opt.as_mut().ok_or_else(|| {
            JsValue::from_str("Rich text rendering requires backend fonts to be registered")
        })?;
        let (mut paragraphs, mut layout_height) =
            text_layout::build_cosmic_paragraphs(font_system, txt, scale);

        if body.map(|value| value.auto_fit.as_str()) == Some("shrink") {
            while layout_height > available_height && scale > 0.2 {
                scale = (scale - 0.05).max(0.2);
                (paragraphs, layout_height) =
                    text_layout::build_cosmic_paragraphs(font_system, txt, scale);
            }
        }

        // spAutoFit grows the text shape when its content needs more height.
        // With vertOverflow=overflow, the shape geometry stays fixed but the
        // text bitmap must still extend far enough to paint overflowing lines.
        let vertical_text = body
            .map(|value| value.text_direction != "horz")
            .unwrap_or(false);
        let allows_vertical_overflow = body
            .map(|value| value.vertical_overflow == "overflow")
            .unwrap_or(true);
        let render_height = if body.map(|value| value.auto_fit.as_str()) == Some("shape")
            || allows_vertical_overflow
        {
            available_height.max(layout_height)
        } else {
            available_height
        };

        #[cfg(debug_assertions)]
        {
            let body_info = body
                .map(|value| {
                    format!(
                        "autoFit={} vertOverflow={} margins=({:.2},{:.2},{:.2},{:.2})",
                        value.auto_fit,
                        value.vertical_overflow,
                        value.margin_left,
                        value.margin_top,
                        value.margin_right,
                        value.margin_bottom
                    )
                })
                .unwrap_or_else(|| "body=default".to_string());
            web_sys::console::log_1(&JsValue::from_str(&format!(
                "[TextBoxLayout] id={} rect=({:.2}x{:.2}) layoutHeight={:.2} renderHeight={:.2} {}",
                txt.id, txt.rect.w, txt.rect.h, layout_height, render_height, body_info
            )));
        }
        let bitmap_height = (render_height * raster_scale).ceil() as u32;

        let vertical_offset = if vertical_text {
            // Vertical columns receive their own anchor offset in the text
            // layout because their measured heights can differ.
            0.0
        } else {
            match body.map(|value| value.vertical_anchor.as_str()) {
                // OOXML vertical anchoring centers the line box even when the
                // noAutofit shape is a little shorter than the natural metrics.
                // Clamping this remainder to zero makes centered titles drift
                // toward the bottom of their header bars.
                Some("middle") => (render_height - layout_height) / 2.0,
                Some("bottom") => render_height - layout_height,
                _ => 0.0,
            }
        };

        let mut pixels = vec![0_u8; bitmap_width as usize * bitmap_height as usize * 4];
        for paragraph in &mut paragraphs {
            let top = paragraph.top + vertical_offset;
            font_renderer::rasterize_buffer(
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
                paragraph.subsequent_line_offset,
                paragraph.hanging_punctuation,
                vertical_text,
                &paragraph.vertical_line_offsets,
                paragraph.vertical_column_height,
                paragraph.column_width,
                &paragraph.font_alignment,
                paragraph.font_size,
                paragraph.metric_line_height,
                paragraph.line_height,
                paragraph.font_metrics,
            );
            if let Some(bullet_buffer) = &paragraph.bullet_buffer {
                font_renderer::rasterize_buffer(
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
                    0.0,
                    false,
                    false,
                    &[],
                    0.0,
                    0.0,
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
            render_height as f64,
        )?;

        if let Some(reflection) = Self::text_reflection(txt) {
            // Run effects apply to rendered glyphs, not to the text line box.
            // Cropping transparent leading/margins keeps a bottom-aligned
            // reflection attached to the actual glyph boundary.
            let Some((source_top, source_bottom)) =
                font_renderer::alpha_vertical_bounds(&pixels, bitmap_width, bitmap_height)
            else {
                return Ok(());
            };
            let source_height = source_bottom - source_top;
            let row_stride = bitmap_width as usize * 4;
            let source_start = source_top as usize * row_stride;
            let source_end = source_bottom as usize * row_stride;
            let reflection_height = (source_height as f32 * reflection.scale_y.abs().max(0.01))
                .ceil()
                .max(1.0) as u32;
            let reflection_pixels = font_renderer::build_reflection_bitmap(
                &pixels[source_start..source_end],
                bitmap_width,
                source_height,
                reflection_height,
                reflection.scale_y,
                reflection.start_alpha,
                reflection.end_alpha,
                reflection.end_position,
                reflection.blur_radius * raster_scale,
            );
            let reflection_canvas: HtmlCanvasElement =
                document.create_element("canvas")?.dyn_into()?;
            reflection_canvas.set_width(bitmap_width);
            reflection_canvas.set_height(reflection_height);
            let reflection_ctx: CanvasRenderingContext2d = reflection_canvas
                .get_context("2d")?
                .ok_or_else(|| JsValue::from_str("Could not create reflection canvas context"))?
                .dyn_into()?;
            let reflection_image = ImageData::new_with_u8_clamped_array_and_sh(
                Clamped(&reflection_pixels),
                bitmap_width,
                reflection_height,
            )?;
            reflection_ctx.put_image_data(&reflection_image, 0.0, 0.0)?;
            let direction = reflection.direction.to_radians();
            let reflection_x = txt.rect.x + reflection.distance * direction.cos();
            let reflection_y = txt.rect.y
                + source_bottom as f32 / raster_scale
                + reflection.distance * direction.sin();
            self.ctx.draw_image_with_html_canvas_element_and_dw_and_dh(
                &reflection_canvas,
                reflection_x as f64,
                reflection_y as f64,
                txt.rect.w as f64,
                reflection_height as f64 / raster_scale as f64,
            )?;
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::font_renderer;
    use super::text_layout;

    #[test]
    fn parses_text_colors_for_rasterization() {
        let solid = text_layout::parse_text_color("#c00000");
        assert_eq!(
            (solid.r(), solid.g(), solid.b(), solid.a()),
            (192, 0, 0, 255)
        );

        let translucent = text_layout::parse_text_color("rgba(1,2,3,0.5)");
        assert_eq!(
            (translucent.r(), translucent.g(), translucent.b()),
            (1, 2, 3)
        );
        assert!((126..=128).contains(&translucent.a()));
    }

    #[test]
    fn blends_rasterized_glyph_pixels() {
        let mut target = [0_u8; 4];
        font_renderer::blend_pixel(&mut target, 0, cosmic_text::Color::rgba(10, 20, 30, 128));
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
                    let has_effect = shp
                        .computed_style
                        .as_ref()
                        .and_then(|style| style.effects.as_ref())
                        .map(|effects| effects.outer_shadow.is_some() || effects.glow.is_some())
                        .unwrap_or(false);

                    if has_effect {
                        let transform = self.ctx.get_transform()?;
                        let device_scale = ((transform.a() * transform.a()
                            + transform.b() * transform.b())
                        .sqrt() as f32)
                            .max(0.1);
                        // The custom mask contains only shadow pixels, so sx/sy can change
                        // the shadow silhouette without exposing a second source outline.
                        effects::render_custom_shape_effect(&self.ctx, shp, device_scale)?;
                    }

                    effects::clear_canvas_shadow(&self.ctx);
                    shape_renderer::paint_shape(&self.ctx, shp);
                    self.ctx.restore();
                }
                Element::Text(txt) => {
                    self.ctx.save();
                    let clips_to_shape = txt
                        .body
                        .as_ref()
                        .map(|body| {
                            body.auto_fit != "shape" && body.vertical_overflow != "overflow"
                        })
                        .unwrap_or(false);
                    if clips_to_shape {
                        self.ctx.begin_path();
                        self.ctx.rect(
                            txt.rect.x as f64,
                            txt.rect.y as f64,
                            txt.rect.w as f64,
                            txt.rect.h as f64,
                        );
                        self.ctx.clip();
                    }
                    let normalized = text_layout::normalize_text_element(txt);
                    self.render_rich_text(&normalized)?;
                    self.ctx.restore();
                }
                Element::Image(img) => {
                    image_renderer::render_image(&self.ctx, img, images_obj)?;
                }
            }
        }
        Ok(())
    }
}
