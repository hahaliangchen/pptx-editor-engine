mod ast;

use wasm_bindgen::prelude::*;
use wasm_bindgen::JsCast;
use web_sys::CanvasRenderingContext2d;
use ast::{Slide, Element};
use cosmic_text::{FontSystem, Buffer, Metrics, Attrs, Family, Weight, Shaping};
use cosmic_text::fontdb::Source;
use std::sync::Arc;

#[wasm_bindgen]
pub struct RustPptRenderer {
    ctx: CanvasRenderingContext2d,
    font_system: Option<FontSystem>,
    registered_fonts: Vec<Arc<Vec<u8>>>,
}

#[wasm_bindgen]
impl RustPptRenderer {
    #[wasm_bindgen(constructor)]
    pub fn new(ctx: CanvasRenderingContext2d) -> Self {
        std::panic::set_hook(Box::new(console_error_panic_hook::hook));
        Self {
            ctx,
            font_system: None,
            registered_fonts: Vec::new(),
        }
    }

    #[wasm_bindgen]
    pub fn register_font(&mut self, font_bytes: &[u8]) {
        let arc_bytes = Arc::new(font_bytes.to_vec());
        self.registered_fonts.push(arc_bytes);
        
        let sources = self.registered_fonts.iter().map(|f| {
            Source::Binary(f.clone())
        }).collect::<Vec<_>>();
        
        // Re-create the font system with the registered fonts
        let mut fs = FontSystem::new_with_fonts(sources.into_iter());
        fs.db_mut().set_sans_serif_family("Roboto");
        fs.db_mut().set_serif_family("Roboto");
        fs.db_mut().set_monospace_family("Roboto");

        // Log loaded font faces to the browser console for debugging
        for face in fs.db().faces() {
            let families_str = face.families.iter().map(|f| &f.0).cloned().collect::<Vec<_>>().join(", ");
            let msg = format!(
                "Loaded Font Face: Families=[{}], Weight={:?}, Style={:?}",
                families_str, face.weight, face.style
            );
            let _ = web_sys::console::log_1(&JsValue::from_str(&msg));
        }

        self.font_system = Some(fs);
    }

    // Render single slide from JSON AST, with images_obj mapping URLs to HTMLImageElements
    #[wasm_bindgen]
    pub fn render_slide(&mut self, slide_json: &str, images_obj: &JsValue) -> Result<(), JsValue> {
        // Debug: Log all currently registered fonts in the WASM font database
        if let Some(fs) = &self.font_system {
            let count = fs.db().faces().count();
            let mut face_names = Vec::new();
            for face in fs.db().faces() {
                let fams = face.families.iter().map(|f| &f.0).cloned().collect::<Vec<_>>().join(",");
                face_names.push(format!("{} (weight: {:?})", fams, face.weight));
            }
            let msg = format!(
                "[WASM Debug] render_slide called. Active font_system exists. Total faces in database: {}. Faces: {:?}",
                count, face_names
            );
            let _ = web_sys::console::log_1(&JsValue::from_str(&msg));
        } else {
            let _ = web_sys::console::log_1(&JsValue::from_str(
                "[WASM Debug] render_slide called. self.font_system is NONE (no fonts loaded yet)."
            ));
        }

        let slide: Slide = serde_json::from_str(slide_json)
            .map_err(|e| JsValue::from_str(&format!("JSON Parse Error: {}", e)))?;

        // 1. Clear and fill background (reset transform to cover physical canvas size)
        self.ctx.save();
        let _ = self.ctx.set_transform(1.0, 0.0, 0.0, 1.0, 0.0, 0.0);
        if let Some(canvas) = self.ctx.canvas() {
            self.ctx.set_fill_style_str("#ffffff");
            self.ctx.fill_rect(0.0, 0.0, canvas.width() as f64, canvas.height() as f64);
        }
        self.ctx.restore();

        // 2. Render elements from bottom to top (Z-Index order)
        for element in &slide.elements {
            match element {
                Element::Shape(shp) => {
                    self.ctx.save();
                    self.ctx.set_fill_style_str(&shp.fill);
                    if shp.shape_type == "rect" {
                        self.ctx.fill_rect(
                            shp.rect.x as f64,
                            shp.rect.y as f64,
                            shp.rect.w as f64,
                            shp.rect.h as f64,
                        );
                    } else if shp.shape_type == "ellipse" {
                        self.ctx.begin_path();
                        let _ = self.ctx.ellipse(
                            (shp.rect.x + shp.rect.w / 2.0) as f64,
                            (shp.rect.y + shp.rect.h / 2.0) as f64,
                            (shp.rect.w / 2.0) as f64,
                            (shp.rect.h / 2.0) as f64,
                            0.0,
                            0.0,
                            2.0 * std::f64::consts::PI,
                        );
                        self.ctx.fill();
                    } else if shp.shape_type == "triangle" {
                        self.ctx.begin_path();
                        self.ctx.move_to((shp.rect.x + shp.rect.w / 2.0) as f64, shp.rect.y as f64);
                        self.ctx.line_to((shp.rect.x + shp.rect.w) as f64, (shp.rect.y + shp.rect.h) as f64);
                        self.ctx.line_to(shp.rect.x as f64, (shp.rect.y + shp.rect.h) as f64);
                        self.ctx.close_path();
                        self.ctx.fill();
                    }
                    self.ctx.restore();
                }
                Element::Text(txt) => {
                    self.ctx.save();
                    self.ctx.set_fill_style_str(&txt.style.color);
                    
                    let font_weight = if txt.style.bold { "bold" } else { "normal" };
                    
                    // Set Canvas font parameters
                    self.ctx.set_font(&format!("{} {}px sans-serif", font_weight, txt.style.font_size));
                    
                    // Alignments
                    let text_align = match txt.style.align.as_str() {
                        "center" => "center",
                        "right" => "right",
                        _ => "left",
                    };
                    self.ctx.set_text_align(text_align);
                    self.ctx.set_text_baseline("top");

                    // Check if the text contains any script characters that Roboto doesn't support
                    // (like CJK, Arabic, Hebrew, Emojis, etc., which are >= 0x0530 in Unicode).
                    // Since cosmic-text lacks fallback fonts for these scripts in our WASM sandbox,
                    // it will panic with 'no default font found'. We fall back to native Canvas.
                    let has_non_latin = txt.content.chars().any(|c| (c as u32) >= 0x0530);

                    // Safety check: If no fonts are registered, font system is not ready, or text contains non-latin scripts,
                    // fall back gracefully to standard canvas text rendering.
                    let font_system = if has_non_latin {
                        None
                    } else {
                        match &mut self.font_system {
                            Some(fs) => {
                                // Verify that the query for default SansSerif succeeds.
                                // If it fails, cosmic-text will panic internally on text shaping,
                                // so we force a graceful fallback to browser canvas layout.
                                let query = cosmic_text::fontdb::Query {
                                    families: &[cosmic_text::fontdb::Family::SansSerif],
                                    ..Default::default()
                                };
                                if fs.db().query(&query).is_none() {
                                    None
                                } else {
                                    Some(fs)
                                }
                            }
                            None => None,
                        }
                    };

                    let font_system = match font_system {
                        Some(fs) => fs,
                        None => {
                            if has_non_latin {
                                let _ = web_sys::console::log_1(&JsValue::from_str(
                                    "[WASM Info] Text contains non-Latin/CJK characters. Using native browser Canvas renderer for system font support."
                                ));
                            } else {
                                let _ = web_sys::console::warn_1(&JsValue::from_str(
                                    "WASM Font Database is empty or lacks a valid SansSerif fallback font. Fallback to standard Canvas text layout."
                                ));
                            }
                            let lines: Vec<&str> = txt.content.split('\n').collect();
                            let line_height = (txt.style.font_size * 1.25) as f64;
                            let start_y = txt.rect.y as f64 + 4.0;
                            for (i, line) in lines.iter().enumerate() {
                                let y = start_y + (i as f64) * line_height;
                                let x = match text_align {
                                    "center" => (txt.rect.x + txt.rect.w / 2.0) as f64,
                                    "right" => (txt.rect.x + txt.rect.w) as f64,
                                    _ => txt.rect.x as f64 + 8.0,
                                };
                                let _ = self.ctx.fill_text(line, x, y);
                            }
                            self.ctx.restore();
                            continue;
                        }
                    };

                    // 2a. Perform text layout using cosmic-text
                    let font_size = txt.style.font_size;
                    let line_height = font_size * 1.25;
                    let mut buffer = Buffer::new(font_system, Metrics::new(font_size, line_height));

                    let pad_x = 8.0;
                    let pad_y = 4.0;
                    let wrap_w = txt.rect.w - pad_x * 2.0;
                    let wrap_h = txt.rect.h - pad_y * 2.0;

                    // Set layout constraints
                    buffer.set_size(font_system, wrap_w, wrap_h);

                    // Use Roboto (which we register) for accurate metrics
                    let mut attrs = Attrs::new().family(Family::Name("Roboto"));
                    if txt.style.bold {
                        attrs = attrs.weight(Weight::BOLD);
                    }

                    // Feed content & layout
                    buffer.set_text(font_system, &txt.content, attrs, Shaping::Advanced);
                    buffer.shape_until_scroll(font_system);

                    // 2b. Draw wrapped layout lines
                    let start_x = txt.rect.x as f64 + pad_x as f64;
                    let start_y = txt.rect.y as f64 + pad_y as f64;

                    for run in buffer.layout_runs() {
                        let paragraph_text = &buffer.lines[run.line_i].text();
                        
                        // Extract run substring via glyph boundaries
                        let mut min_index = paragraph_text.len();
                        let mut max_index = 0;
                        for glyph in run.glyphs {
                            let start = glyph.start;
                            let end = glyph.end;
                            if start < min_index { min_index = start; }
                            if end > max_index { max_index = end; }
                        }
                        
                        let line_text = if min_index < max_index {
                            &paragraph_text[min_index..max_index]
                        } else {
                            ""
                        };

                        let y = start_y + run.line_y as f64;
                        let x = match text_align {
                            "center" => start_x + (wrap_w / 2.0) as f64,
                            "right" => start_x + wrap_w as f64,
                            _ => start_x,
                        };

                        let _ = self.ctx.fill_text(line_text, x, y);
                    }
                    self.ctx.restore();
                }
                Element::Image(img) => {
                    self.ctx.save();
                    if let Some(img_val) = js_sys::Reflect::get(images_obj, &JsValue::from_str(&img.url)).ok() {
                        if !img_val.is_undefined() && !img_val.is_null() {
                            if let Ok(html_img) = img_val.dyn_into::<web_sys::HtmlImageElement>() {
                                let _ = self.ctx.draw_image_with_html_image_element_and_dw_and_dh(
                                    &html_img,
                                    img.rect.x as f64,
                                    img.rect.y as f64,
                                    img.rect.w as f64,
                                    img.rect.h as f64,
                                );
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
