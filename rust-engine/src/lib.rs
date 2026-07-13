mod ast;

use ast::{Element, Slide, TextElement, TextStyle};
use cosmic_text::fontdb::Source;
use cosmic_text::{Attrs, Buffer, Family, FontSystem, Metrics, Shaping, Style, Weight};
use std::sync::Arc;
use wasm_bindgen::prelude::*;
use wasm_bindgen::JsCast;
use web_sys::CanvasRenderingContext2d;

#[wasm_bindgen]
pub struct RustPptRenderer {
    ctx: CanvasRenderingContext2d,
    font_system: Option<FontSystem>,
    registered_fonts: Vec<Arc<Vec<u8>>>,
}

#[derive(Clone)]
struct RichFragment {
    text: String,
    style: TextStyle,
    width: f32,
}

struct RichLine {
    fragments: Vec<RichFragment>,
    width: f32,
    height: f32,
    x: f32,
    y: f32,
    bullet: Option<(String, TextStyle, f32)>,
}

struct RichLayout {
    lines: Vec<RichLine>,
    height: f32,
    max_width: f32,
}

impl RustPptRenderer {
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

    fn layout_tokens(text: &str) -> Vec<String> {
        let mut tokens = Vec::new();
        let mut current = String::new();
        let mut current_kind: Option<u8> = None;

        for ch in text.chars() {
            if ch == '\n' {
                if !current.is_empty() {
                    tokens.push(std::mem::take(&mut current));
                }
                tokens.push("\n".to_string());
                current_kind = None;
                continue;
            }

            let kind = if ch.is_whitespace() {
                1
            } else if Self::is_east_asian_text(&ch.to_string()) {
                2
            } else {
                3
            };
            if kind == 2 {
                if !current.is_empty() {
                    tokens.push(std::mem::take(&mut current));
                }
                tokens.push(ch.to_string());
                current_kind = None;
            } else {
                if current_kind.is_some() && current_kind != Some(kind) && !current.is_empty() {
                    tokens.push(std::mem::take(&mut current));
                }
                current.push(ch);
                current_kind = Some(kind);
            }
        }
        if !current.is_empty() {
            tokens.push(current);
        }
        tokens
    }

    fn set_text_font(&self, style: &TextStyle, text: &str, scale: f32) {
        let font_style = if style.italic { "italic" } else { "normal" };
        let font_weight = if style.bold { "bold" } else { "normal" };
        let requested_family =
            if Self::is_east_asian_text(text) && !style.east_asian_font_family.is_empty() {
                &style.east_asian_font_family
            } else {
                &style.font_family
            };
        let family = requested_family.replace('\'', "");
        self.ctx.set_font(&format!(
            "{} {} {}px '{}'",
            font_style,
            font_weight,
            style.font_size * scale,
            family
        ));
    }

    fn measure_fragment(&self, text: &str, style: &TextStyle, scale: f32) -> f32 {
        self.set_text_font(style, text, scale);
        self.ctx
            .measure_text(text)
            .map(|metrics| metrics.width() as f32)
            .unwrap_or_else(|_| text.chars().count() as f32 * style.font_size * scale * 0.55)
    }

    fn build_rich_layout(&self, txt: &TextElement, scale: f32) -> RichLayout {
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
        let inner_left = txt.rect.x + margin_left;
        let inner_width = (txt.rect.w - margin_left - margin_right).max(1.0);
        let mut positioned_lines = Vec::new();
        let mut cursor_y = margin_top;
        let mut max_width: f32 = 0.0;

        for paragraph in &txt.paragraphs {
            cursor_y += paragraph.style.space_before * scale;
            let first_style = paragraph
                .runs
                .first()
                .map(|run| run.style.clone())
                .unwrap_or_else(|| txt.style.clone());
            let paragraph_left = paragraph.style.margin_left.max(0.0);
            let available_width = (inner_width - paragraph_left).max(1.0);
            let first_line_indent = if paragraph.bullet.is_none() {
                paragraph.style.indent
            } else {
                0.0
            };

            let mut lines: Vec<RichLine> = Vec::new();
            let mut current = RichLine {
                fragments: Vec::new(),
                width: 0.0,
                height: 0.0,
                x: 0.0,
                y: 0.0,
                bullet: None,
            };

            for run in &paragraph.runs {
                for token in Self::layout_tokens(&run.content) {
                    if token == "\n" {
                        lines.push(current);
                        current = RichLine {
                            fragments: Vec::new(),
                            width: 0.0,
                            height: 0.0,
                            x: 0.0,
                            y: 0.0,
                            bullet: None,
                        };
                        continue;
                    }

                    let current_indent = if lines.is_empty() {
                        first_line_indent.max(0.0)
                    } else {
                        0.0
                    };
                    let current_width_limit = (available_width - current_indent).max(1.0);
                    let token_width = self.measure_fragment(&token, &run.style, scale);
                    if token.trim().is_empty()
                        && current.width > 0.0
                        && current.width + token_width > current_width_limit
                    {
                        lines.push(current);
                        current = RichLine {
                            fragments: Vec::new(),
                            width: 0.0,
                            height: 0.0,
                            x: 0.0,
                            y: 0.0,
                            bullet: None,
                        };
                        continue;
                    }
                    if !token.trim().is_empty()
                        && current.width > 0.0
                        && current.width + token_width > current_width_limit
                    {
                        lines.push(current);
                        current = RichLine {
                            fragments: Vec::new(),
                            width: 0.0,
                            height: 0.0,
                            x: 0.0,
                            y: 0.0,
                            bullet: None,
                        };
                    }
                    if token.trim().is_empty() && current.width == 0.0 {
                        continue;
                    }

                    for ch in token.chars() {
                        let mut candidate_width =
                            self.measure_fragment(&ch.to_string(), &run.style, scale);
                        let mut extends_last = false;
                        if let Some(last) = current.fragments.last() {
                            if last.style == run.style {
                                let candidate = format!("{}{}", last.text, ch);
                                candidate_width =
                                    self.measure_fragment(&candidate, &run.style, scale);
                                extends_last = true;
                            }
                        }
                        let next_width = if extends_last {
                            current.width - current.fragments.last().map(|f| f.width).unwrap_or(0.0)
                                + candidate_width
                        } else {
                            current.width + candidate_width
                        };

                        let current_indent = if lines.is_empty() {
                            first_line_indent.max(0.0)
                        } else {
                            0.0
                        };
                        let current_width_limit = (available_width - current_indent).max(1.0);
                        if next_width > current_width_limit && current.width > 0.0 {
                            lines.push(current);
                            current = RichLine {
                                fragments: Vec::new(),
                                width: 0.0,
                                height: 0.0,
                                x: 0.0,
                                y: 0.0,
                                bullet: None,
                            };
                            candidate_width =
                                self.measure_fragment(&ch.to_string(), &run.style, scale);
                            extends_last = false;
                        }

                        if extends_last {
                            if let Some(last) = current.fragments.last_mut() {
                                last.text.push(ch);
                                current.width = current.width - last.width + candidate_width;
                                last.width = candidate_width;
                            }
                        } else {
                            current.fragments.push(RichFragment {
                                text: ch.to_string(),
                                style: run.style.clone(),
                                width: candidate_width,
                            });
                            current.width += candidate_width;
                        }
                    }
                }
            }
            lines.push(current);

            for (line_index, mut line) in lines.into_iter().enumerate() {
                let max_font_size = line
                    .fragments
                    .iter()
                    .map(|fragment| fragment.style.font_size * scale)
                    .fold(first_style.font_size * scale, f32::max);
                line.height = paragraph
                    .style
                    .line_spacing
                    .as_ref()
                    .map(|spacing| {
                        if spacing.unit == "points" {
                            (spacing.value * scale).max(max_font_size)
                        } else {
                            (max_font_size * spacing.value).max(max_font_size)
                        }
                    })
                    .unwrap_or(max_font_size * 1.2);

                let indent = if line_index == 0 {
                    first_line_indent
                } else {
                    0.0
                };
                let line_left = inner_left + paragraph_left + indent;
                let line_width = (available_width - indent.max(0.0)).max(1.0);
                line.x = match paragraph.style.align.as_str() {
                    "center" => line_left + (line_width - line.width).max(0.0) / 2.0,
                    "right" => line_left + (line_width - line.width).max(0.0),
                    _ => line_left,
                };
                line.y = cursor_y;
                if line_index == 0 {
                    if let Some(bullet) = &paragraph.bullet {
                        let mut bullet_style = first_style.clone();
                        bullet_style.color = bullet.color.clone();
                        line.bullet = Some((
                            bullet.char.clone(),
                            bullet_style,
                            inner_left + paragraph_left + paragraph.style.indent,
                        ));
                    }
                }
                max_width = max_width.max(line.width);
                cursor_y += line.height;
                positioned_lines.push(line);
            }
            cursor_y += paragraph.style.space_after * scale;
        }

        RichLayout {
            lines: positioned_lines,
            height: cursor_y + margin_bottom,
            max_width,
        }
    }

    fn render_rich_text(&self, txt: &TextElement) {
        let body = txt.body.as_ref();
        let available_height = txt.rect.h.max(1.0);
        let mut scale = body
            .map(|value| value.font_scale)
            .unwrap_or(1.0)
            .clamp(0.2, 1.0);
        let mut layout = self.build_rich_layout(txt, scale);

        if body.map(|value| value.auto_fit.as_str()) == Some("shrink") {
            while (layout.height > available_height || layout.max_width > txt.rect.w) && scale > 0.2
            {
                scale = (scale - 0.05).max(0.2);
                layout = self.build_rich_layout(txt, scale);
            }
        }

        let vertical_offset = match body.map(|value| value.vertical_anchor.as_str()) {
            Some("middle") => (available_height - layout.height).max(0.0) / 2.0,
            Some("bottom") => (available_height - layout.height).max(0.0),
            _ => 0.0,
        };

        self.ctx.set_text_align("left");
        self.ctx.set_text_baseline("top");
        for line in &layout.lines {
            let y = txt.rect.y + line.y + vertical_offset;
            if let Some((bullet, style, x)) = &line.bullet {
                self.set_text_font(style, bullet, scale);
                self.ctx.set_fill_style_str(&style.color);
                let _ = self.ctx.fill_text(bullet, *x as f64, y as f64);
            }

            let mut x = line.x;
            for fragment in &line.fragments {
                self.set_text_font(&fragment.style, &fragment.text, scale);
                self.ctx.set_fill_style_str(&fragment.style.color);
                let _ = self.ctx.fill_text(&fragment.text, x as f64, y as f64);
                x += fragment.width;
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::RustPptRenderer;

    #[test]
    fn layout_tokens_keep_words_and_split_east_asian_text() {
        assert_eq!(
            RustPptRenderer::layout_tokens("Hello world中文\nnext"),
            vec!["Hello", " ", "world", "中", "文", "\n", "next"]
        );
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
                    let is_filled = shp.fill != "transparent";
                    if is_filled {
                        self.ctx.set_fill_style_str(&shp.fill);
                    }

                    if shp.shape_type == "rect" {
                        if is_filled {
                            self.ctx.fill_rect(
                                shp.rect.x as f64,
                                shp.rect.y as f64,
                                shp.rect.w as f64,
                                shp.rect.h as f64,
                            );
                        }
                        if let Some(border) = &shp.border {
                            self.ctx.set_stroke_style_str(&border.color);
                            self.ctx.set_line_width(border.width as f64);
                            self.ctx.stroke_rect(
                                shp.rect.x as f64,
                                shp.rect.y as f64,
                                shp.rect.w as f64,
                                shp.rect.h as f64,
                            );
                        }
                    } else if shp.shape_type == "roundRect" {
                        let x = shp.rect.x as f64;
                        let y = shp.rect.y as f64;
                        let w = shp.rect.w as f64;
                        let h = shp.rect.h as f64;
                        let radius = (w.min(h) * 0.13).max(1.0);
                        self.ctx.begin_path();
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
                        self.ctx.close_path();
                        if is_filled {
                            self.ctx.fill();
                        }
                        if let Some(border) = &shp.border {
                            self.ctx.set_stroke_style_str(&border.color);
                            self.ctx.set_line_width(border.width as f64);
                            self.ctx.stroke();
                        }
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
                        if is_filled {
                            self.ctx.fill();
                        }
                        if let Some(border) = &shp.border {
                            self.ctx.set_stroke_style_str(&border.color);
                            self.ctx.set_line_width(border.width as f64);
                            self.ctx.stroke();
                        }
                    } else if shp.shape_type == "triangle" {
                        self.ctx.begin_path();
                        self.ctx
                            .move_to((shp.rect.x + shp.rect.w / 2.0) as f64, shp.rect.y as f64);
                        self.ctx.line_to(
                            (shp.rect.x + shp.rect.w) as f64,
                            (shp.rect.y + shp.rect.h) as f64,
                        );
                        self.ctx
                            .line_to(shp.rect.x as f64, (shp.rect.y + shp.rect.h) as f64);
                        self.ctx.close_path();
                        if is_filled {
                            self.ctx.fill();
                        }
                        if let Some(border) = &shp.border {
                            self.ctx.set_stroke_style_str(&border.color);
                            self.ctx.set_line_width(border.width as f64);
                            self.ctx.stroke();
                        }
                    }
                    self.ctx.restore();
                }
                Element::Text(txt) => {
                    self.ctx.save();
                    if !txt.paragraphs.is_empty() {
                        self.ctx.begin_path();
                        self.ctx.rect(
                            txt.rect.x as f64,
                            txt.rect.y as f64,
                            txt.rect.w as f64,
                            txt.rect.h as f64,
                        );
                        self.ctx.clip();
                        self.render_rich_text(txt);
                        self.ctx.restore();
                        continue;
                    }
                    self.ctx.set_fill_style_str(&txt.style.color);

                    let font_weight = if txt.style.bold { "bold" } else { "normal" };

                    // Set Canvas font parameters
                    let font_style = if txt.style.italic { "italic" } else { "normal" };
                    self.ctx.set_font(&format!(
                        "{} {} {}px {}",
                        font_style, font_weight, txt.style.font_size, txt.style.font_family
                    ));

                    // Alignments
                    let text_align = match txt.style.align.as_str() {
                        "center" => "center",
                        "right" => "right",
                        _ => "left",
                    };
                    self.ctx.set_text_align(text_align);
                    self.ctx.set_text_baseline("top");

                    // Safety check: Verify that the font database is initialized and has at least one font queryable.
                    // If it is empty, we fall back gracefully to standard Canvas text wrapping.
                    let font_system = match &mut self.font_system {
                        Some(fs) => {
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
                    };

                    let font_system = match font_system {
                        Some(fs) => fs,
                        None => {
                            let _ = web_sys::console::warn_1(&JsValue::from_str(
                                "[WASM Warning] WASM Font Database is empty or lacks a valid SansSerif fallback font. Falling back to native browser Canvas renderer with auto-wrap."
                            ));

                            // Native canvas wrapping layout
                            let pad_x = if txt.body.is_some() { 0.0 } else { 8.0 };
                            let max_width = (txt.rect.w - pad_x * 2.0) as f64;
                            let mut wrapped_lines = Vec::new();
                            let raw_lines = txt.content.split('\n');

                            for raw_line in raw_lines {
                                if raw_line.is_empty() {
                                    wrapped_lines.push("".to_string());
                                    continue;
                                }
                                let mut current_line = String::new();
                                for c in raw_line.chars() {
                                    let test_line = if current_line.is_empty() {
                                        c.to_string()
                                    } else {
                                        format!("{}{}", current_line, c)
                                    };
                                    if let Ok(metrics) = self.ctx.measure_text(&test_line) {
                                        if metrics.width() > max_width && !current_line.is_empty() {
                                            wrapped_lines.push(current_line);
                                            current_line = c.to_string();
                                        } else {
                                            current_line = test_line;
                                        }
                                    } else {
                                        current_line = test_line;
                                    }
                                }
                                if !current_line.is_empty() {
                                    wrapped_lines.push(current_line);
                                }
                            }

                            let line_height = (txt.style.font_size * 1.25) as f64;
                            let pad_y = if txt.body.is_some() { 0.0 } else { 4.0 };
                            let start_y = txt.rect.y as f64 + pad_y;
                            for (i, line) in wrapped_lines.iter().enumerate() {
                                let y = start_y + (i as f64) * line_height;
                                let x = match text_align {
                                    "center" => (txt.rect.x + txt.rect.w / 2.0) as f64,
                                    "right" => (txt.rect.x + txt.rect.w) as f64,
                                    _ => txt.rect.x as f64 + pad_x as f64,
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

                    let pad_x = if txt.body.is_some() { 0.0 } else { 8.0 };
                    let pad_y = if txt.body.is_some() { 0.0 } else { 4.0 };
                    let wrap_w = txt.rect.w - pad_x * 2.0;
                    let wrap_h = txt.rect.h - pad_y * 2.0;

                    // Set layout constraints
                    buffer.set_size(font_system, wrap_w, wrap_h);

                    // Shape with the exact family supplied by the backend-resolved AST.
                    let mut attrs = Attrs::new().family(Family::Name(&txt.style.font_family));
                    if txt.style.bold {
                        attrs = attrs.weight(Weight::BOLD);
                    }
                    if txt.style.italic {
                        attrs = attrs.style(Style::Italic);
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
                            if start < min_index {
                                min_index = start;
                            }
                            if end > max_index {
                                max_index = end;
                            }
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
