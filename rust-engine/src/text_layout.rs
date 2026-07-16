use crate::ast::{
    ParagraphStyle, TextBodyProperties, TextElement, TextParagraph, TextRun, TextStyle,
};
use cosmic_text::fontdb::{self, Query};
use cosmic_text::{
    Align, Attrs, AttrsList, Buffer, BufferLine, CacheKeyFlags, Color, Family, FontSystem, Hinting,
    LineEnding, Metrics, Shaping, Style, Weight, Wrap,
};

const WPS_FONT_METRIC_SCALE: f32 = 0.92;
// WPS applies the explicit percentage increment a little more tightly than
// the raw DrawingML percentage when it is combined with the font metric box.
const WPS_PERCENT_SPACING_SCALE: f32 = 0.8;
pub struct CosmicParagraph {
    pub buffer: Buffer,
    pub bullet_buffer: Option<Buffer>,
    pub x: f32,
    pub column_width: f32,
    pub vertical_line_offsets: Vec<f32>,
    pub vertical_column_height: f32,
    pub first_line_offset: f32,
    pub subsequent_line_offset: f32,
    pub bullet_x: f32,
    pub top: f32,
    pub hanging_punctuation: bool,
    pub font_alignment: String,
    pub font_size: f32,
    pub metric_line_height: f32,
    pub line_height: f32,
    pub font_metrics: Option<FontMetricSample>,
}

#[derive(Clone, Copy)]
pub struct FontMetricSample {
    pub ascent: f32,
    pub descent: f32,
    pub leading: f32,
    pub line_height: f32,
}

pub fn is_east_asian_text(text: &str) -> bool {
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

pub fn is_vertical_horizontal_char(ch: char) -> bool {
    ch.is_ascii_alphanumeric()
}

pub fn is_vertical_rotated_char(ch: char) -> bool {
    ch.is_ascii() && !ch.is_ascii_control()
}

pub fn is_vertical_compact_char(ch: char) -> bool {
    !is_vertical_fullwidth_char(ch)
        && (is_vertical_rotated_char(ch) || matches!(ch, '\u{2000}'..='\u{206F}'))
}

pub fn is_vertical_fullwidth_char(ch: char) -> bool {
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
            | 0x3000..=0x303F
            | 0xFF01..=0xFF65
    )
}

pub fn parse_text_color(value: &str) -> Color {
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

pub fn cosmic_attrs<'a>(
    style: &'a TextStyle,
    text: &str,
    scale: f32,
    line_height: f32,
) -> Attrs<'a> {
    let requested_family = if is_east_asian_text(text) && !style.east_asian_font_family.is_empty() {
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
        .color(parse_text_color(&style.color))
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

pub fn font_metric_line_height(
    font_system: &mut FontSystem,
    style: &TextStyle,
    text: &str,
    scale: f32,
) -> Option<FontMetricSample> {
    let font_size = (style.font_size * scale).max(1.0);
    let attrs = cosmic_attrs(style, text, scale, font_size);
    let family = match attrs.family {
        Family::Name(name) => fontdb::Family::Name(name),
        Family::Serif => fontdb::Family::Serif,
        Family::SansSerif => fontdb::Family::SansSerif,
        Family::Cursive => fontdb::Family::Cursive,
        Family::Fantasy => fontdb::Family::Fantasy,
        Family::Monospace => fontdb::Family::Monospace,
    };
    let families = [family];
    let id = font_system.db().query(&Query {
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

pub fn normalize_text_element(txt: &TextElement) -> TextElement {
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
                vertical_overflow: "overflow".to_string(),
                horizontal_overflow: "overflow".to_string(),
                font_scale: 1.0,
                text_direction: "horz".to_string(),
            })
        }),
    }
}

fn vertical_token_advance(
    font_system: &mut FontSystem,
    style: &TextStyle,
    token: &str,
    scale: f32,
    cjk_line_height: f32,
) -> f32 {
    let font_size = (style.font_size * scale).max(1.0);
    let attrs = cosmic_attrs(style, token, scale, cjk_line_height);
    let mut line = BufferLine::new(
        token,
        LineEnding::None,
        AttrsList::new(&attrs),
        Shaping::Advanced,
    );
    line.set_align(Some(Align::Left));
    let mut buffer = Buffer::new_empty(Metrics::new(font_size, cjk_line_height));
    buffer.lines.push(line);
    buffer.set_hinting(Hinting::Disabled);
    buffer.set_size(Some(100_000.0), Some(100_000.0));
    buffer.set_wrap(Wrap::None);
    buffer.shape_until_scroll(font_system, false);
    let measured_width = buffer
        .layout_runs()
        .map(|run| run.line_w)
        .fold(0.0, f32::max);

    if token.chars().all(is_vertical_fullwidth_char) {
        cjk_line_height.max(font_size)
    } else if token.chars().all(is_vertical_compact_char) {
        // The shaped run width already includes the selected font, weight,
        // kerning and letter spacing. Do not replace it with a fixed Latin
        // scale; A, I, W, punctuation and different fonts have different
        // advances.
        measured_width.max(font_size * 0.2)
    } else {
        cjk_line_height.max(font_size)
    }
}

#[derive(Clone)]
struct HorizontalWrapUnit {
    start: usize,
    end: usize,
    style: TextStyle,
}

fn collect_horizontal_wrap_units(paragraph: &TextParagraph) -> Vec<HorizontalWrapUnit> {
    let mut units = Vec::new();
    let mut byte_offset = 0;
    for run in &paragraph.runs {
        let chars: Vec<(usize, char)> = run.content.char_indices().collect();
        let mut index = 0;
        while index < chars.len() {
            let start = chars[index].0;
            let ch = chars[index].1;
            let mut next = index + 1;
            if ch.is_ascii_alphanumeric() {
                while next < chars.len() && chars[next].1.is_ascii_alphanumeric() {
                    next += 1;
                }
            }
            let end = if next < chars.len() {
                chars[next].0
            } else {
                run.content.len()
            };
            units.push(HorizontalWrapUnit {
                start: byte_offset + start,
                end: byte_offset + end,
                style: run.style.clone(),
            });
            index = next;
        }
        byte_offset += run.content.len();
    }
    units
}

fn measure_horizontal_unit(
    font_system: &mut FontSystem,
    unit: &HorizontalWrapUnit,
    text: &str,
    scale: f32,
    line_height: f32,
) -> f32 {
    let content = &text[unit.start..unit.end];
    let attrs = cosmic_attrs(&unit.style, content, scale, line_height);
    let line = BufferLine::new(
        content,
        LineEnding::None,
        AttrsList::new(&attrs),
        Shaping::Advanced,
    );
    let mut buffer = Buffer::new_empty(Metrics::new(
        (unit.style.font_size * scale).max(1.0),
        line_height,
    ));
    buffer.lines.push(line);
    buffer.set_hinting(Hinting::Disabled);
    buffer.set_size(Some(100_000.0), Some(100_000.0));
    buffer.set_wrap(Wrap::None);
    buffer.shape_until_scroll(font_system, false);
    buffer
        .layout_runs()
        .map(|run| run.line_w)
        .fold(0.0, f32::max)
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

fn build_explicit_wrap_ranges(
    font_system: &mut FontSystem,
    paragraph: &TextParagraph,
    text: &str,
    first_line_width: f32,
    subsequent_line_width: f32,
    scale: f32,
    line_height: f32,
    allow_hanging_punctuation: bool,
) -> Vec<(usize, usize)> {
    let units = collect_horizontal_wrap_units(paragraph);
    if units.is_empty() {
        return vec![(0, 0)];
    }

    let mut ranges = Vec::new();
    let mut line_start = units[0].start;
    let mut line_width = 0.0;
    let mut hanging_end = false;

    for unit in &units {
        let content = &text[unit.start..unit.end];
        if content == "\n" {
            ranges.push((line_start, unit.start));
            line_start = unit.end;
            line_width = 0.0;
            hanging_end = false;
            continue;
        }

        if hanging_end {
            ranges.push((line_start, unit.start));
            line_start = unit.start;
            line_width = 0.0;
            hanging_end = false;
        }

        let advance = measure_horizontal_unit(font_system, unit, text, scale, line_height);
        let line_limit = if ranges.is_empty() {
            first_line_width
        } else {
            subsequent_line_width
        };
        let closes_line = content
            .chars()
            .last()
            .is_some_and(is_hanging_closing_punctuation);
        if line_width > 0.0 && line_width + advance > line_limit {
            if allow_hanging_punctuation && closes_line {
                // Keep the punctuation in this line, but do not let its
                // advance decide whether the preceding CJK glyph fits.
                hanging_end = true;
                continue;
            }
            ranges.push((line_start, unit.start));
            line_start = unit.start;
            line_width = 0.0;
        }
        line_width += advance;
    }

    ranges.push((line_start, text.len()));
    ranges
}

fn attrs_for_text_range(
    paragraph: &TextParagraph,
    first_style: &TextStyle,
    text_start: usize,
    text_end: usize,
    scale: f32,
    line_height: f32,
) -> AttrsList {
    let segment = paragraph
        .runs
        .iter()
        .flat_map(|run| run.content.chars())
        .collect::<String>();
    let default_attrs = cosmic_attrs(first_style, &segment, scale, line_height);
    let mut attrs_list = AttrsList::new(&default_attrs);
    let mut run_start = 0;
    for run in &paragraph.runs {
        let run_end = run_start + run.content.len();
        let overlap_start = text_start.max(run_start);
        let overlap_end = text_end.min(run_end);
        if overlap_start < overlap_end {
            let attrs = cosmic_attrs(&run.style, &run.content, scale, line_height);
            attrs_list.add_span(
                (overlap_start - text_start)..(overlap_end - text_start),
                &attrs,
            );
        }
        run_start = run_end;
    }
    attrs_list
}

pub fn build_cosmic_paragraphs(
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
    let vertical_text = txt
        .body
        .as_ref()
        .map(|body| body.text_direction != "horz")
        .unwrap_or(false);
    let mut cursor_y = margin_top;
    let mut vertical_column = 0_usize;
    let mut vertical_content_height = 0.0_f32;
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
        let font_metrics = paragraph
            .runs
            .iter()
            .filter_map(|run| font_metric_line_height(font_system, &run.style, &run.content, scale))
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
                    // DrawingML spcPct supplies the paragraph's additional
                    // spacing relative to the text size. Keep the natural
                    // font alignment box, then add the percentage delta;
                    // multiplying the complete metric box by spcPct makes
                    // mixed 13pt/12pt paragraphs grow too tall.
                    let extra_spacing =
                        (font_size * (spacing.value - 1.0)).max(0.0) * WPS_PERCENT_SPACING_SCALE;
                    metric_line_height + extra_spacing
                }
            })
            .unwrap_or(metric_line_height);
        // Vertical East Asian writing advances by the font em cell rather
        // than the horizontal paragraph line box (which may include 120%
        // paragraph spacing).
        let vertical_line_height = font_size.max(1.0);

        let text: String = paragraph
            .runs
            .iter()
            .map(|run| run.content.as_str())
            .collect();
        let paragraph_left = paragraph.style.margin_left.max(0.0);
        let text_indent = if paragraph.bullet.is_none() {
            paragraph.style.indent
        } else {
            0.0
        };
        let centered_hanging_paragraph = !vertical_text
            && paragraph.bullet.is_none()
            && paragraph_left > 0.0
            && paragraph.style.indent < 0.0
            && matches!(paragraph.style.align.as_str(), "center" | "right");
        let first_line_offset =
            if !vertical_text && paragraph.bullet.is_none() && !centered_hanging_paragraph {
                paragraph.style.indent
            } else {
                0.0
            };
        // With a centered hanging paragraph, the first line uses the full
        // inner width because the negative indent hangs its prefix to the
        // left. Subsequent lines are centered in the narrower area beginning
        // at marL, which shifts them right by half of that margin.
        let subsequent_line_offset = if centered_hanging_paragraph {
            paragraph_left * 0.5
        } else {
            0.0
        };
        let mut buffers: Vec<(Buffer, Vec<f32>, f32)> = Vec::new();
        if vertical_text {
            // eaVert fills a column top-to-bottom and advances right-to-left
            // when the paragraph does not fit. Keep the capacity based on the
            // same line metrics used by the glyph buffer so short labels such
            // as "门户统一" form the same two columns as WPS.
            let mut token_groups: Vec<Vec<(String, TextStyle)>> = Vec::new();
            let mut token_group = Vec::new();
            for run in &paragraph.runs {
                for ch in run.content.chars() {
                    if ch == '\n' {
                        if !token_group.is_empty() {
                            token_groups.push(std::mem::take(&mut token_group));
                        }
                        continue;
                    }
                    if is_vertical_horizontal_char(ch) {
                        if let Some((last, last_style)) = token_group.last_mut() {
                            if last.chars().all(is_vertical_horizontal_char)
                                && last_style == &run.style
                            {
                                last.push(ch);
                            } else {
                                token_group.push((ch.to_string(), run.style.clone()));
                            }
                        } else {
                            token_group.push((ch.to_string(), run.style.clone()));
                        }
                    } else {
                        token_group.push((ch.to_string(), run.style.clone()));
                    }
                }
            }
            if !token_group.is_empty() {
                token_groups.push(token_group);
            }
            if token_groups.is_empty() {
                token_groups.push(Vec::new());
            }

            let available_height = (txt.rect.h - margin_top - margin_bottom)
                .max(vertical_line_height)
                .max(1.0);
            let mut chunks: Vec<Vec<(String, TextStyle, f32)>> = Vec::new();
            for group in token_groups {
                let mut chunk = Vec::new();
                let mut chunk_height = 0.0;
                for (content, style) in group {
                    let advance = vertical_token_advance(
                        font_system,
                        &style,
                        &content,
                        scale,
                        vertical_line_height,
                    );
                    if !chunk.is_empty() && chunk_height + advance > available_height {
                        chunks.push(std::mem::take(&mut chunk));
                        chunk_height = 0.0;
                    }
                    chunk_height += advance;
                    chunk.push((content, style, advance));
                }
                if !chunk.is_empty() {
                    chunks.push(chunk);
                }
            }
            if chunks.is_empty() {
                chunks.push(Vec::new());
            }

            for chunk in chunks {
                let column_width = font_size.max(1.0);
                let mut buffer = Buffer::new_empty(Metrics::new(font_size, vertical_line_height));
                let mut vertical_line_offsets = Vec::new();
                let mut vertical_height = 0.0;
                for (content, style, advance) in chunk {
                    let attrs = AttrsList::new(&cosmic_attrs(
                        &style,
                        &content,
                        scale,
                        vertical_line_height,
                    ));
                    let mut line =
                        BufferLine::new(&content, LineEnding::None, attrs, Shaping::Advanced);
                    line.set_align(Some(Align::Center));
                    vertical_line_offsets.push(vertical_height);
                    vertical_height += advance.max(1.0);
                    buffer.lines.push(line);
                }
                if buffer.lines.is_empty() {
                    let attrs =
                        AttrsList::new(&cosmic_attrs(first_style, "", scale, vertical_line_height));
                    buffer.lines.push(BufferLine::new(
                        "",
                        LineEnding::None,
                        attrs,
                        Shaping::Advanced,
                    ));
                }
                buffer.set_hinting(Hinting::Disabled);
                buffer.set_size(Some(column_width), Some(100_000.0));
                buffer.set_wrap(Wrap::None);
                buffer.shape_until_scroll(font_system, false);
                buffers.push((buffer, vertical_line_offsets, vertical_height));
            }
        } else {
            // A negative indent is a hanging indent: the first line starts
            // to the left of marL and therefore has a wider measure, while
            // subsequent lines use the paragraph margin normally. Positive
            // indents are the inverse. The text frame itself never changes.
            let first_line_width = if paragraph.bullet.is_none() {
                (inner_width - paragraph_left - text_indent).max(1.0)
            } else {
                (inner_width - paragraph_left).max(1.0)
            };
            let subsequent_line_width = (inner_width - paragraph_left).max(1.0);
            let available_width = subsequent_line_width;
            #[cfg(debug_assertions)]
            web_sys::console::log_1(&wasm_bindgen::JsValue::from_str(&format!(
                "[TextMeasure] id={} rectWidth={:.2} innerWidth={:.2} bodyMargins=({:.2},{:.2}) paragraphMargin={:.2} indent={:.2} hangingPunct={} centeredHanging={} firstLineWidth={:.2} subsequentLineWidth={:.2} text={:?}",
                txt.id,
                txt.rect.w,
                inner_width,
                margin_left,
                margin_right,
                paragraph_left,
                paragraph.style.indent,
                paragraph.style.hanging_punctuation,
                centered_hanging_paragraph,
                first_line_width,
                subsequent_line_width,
                text
            )));
            let align = match paragraph.style.align.as_str() {
                "center" => Align::Center,
                "right" => Align::Right,
                _ => Align::Left,
            };
            let mut buffer = Buffer::new_empty(Metrics::new(font_size, line_height));
            if is_east_asian_text(&text)
                && (paragraph.style.hanging_punctuation || paragraph.style.indent != 0.0)
            {
                for (start, end) in build_explicit_wrap_ranges(
                    font_system,
                    paragraph,
                    &text,
                    first_line_width,
                    subsequent_line_width,
                    scale,
                    line_height,
                    paragraph.style.hanging_punctuation,
                ) {
                    let attrs_list = attrs_for_text_range(
                        paragraph,
                        first_style,
                        start,
                        end,
                        scale,
                        line_height,
                    );
                    let mut line = BufferLine::new(
                        &text[start..end],
                        LineEnding::None,
                        attrs_list,
                        Shaping::Advanced,
                    );
                    line.set_align(Some(align));
                    buffer.lines.push(line);
                }
                // The lines were broken explicitly so the terminal closing
                // punctuation can hang without widening every later line.
                buffer.set_wrap(Wrap::None);
            } else {
                let default_attrs = cosmic_attrs(first_style, &text, scale, line_height);
                let mut attrs_list = AttrsList::new(&default_attrs);
                let mut byte_offset = 0;
                for run in &paragraph.runs {
                    let end = byte_offset + run.content.len();
                    let run_attrs = cosmic_attrs(&run.style, &run.content, scale, line_height);
                    attrs_list.add_span(byte_offset..end, &run_attrs);
                    byte_offset = end;
                }
                let mut line =
                    BufferLine::new(&text, LineEnding::None, attrs_list, Shaping::Advanced);
                line.set_align(Some(align));
                buffer.lines.push(line);
                buffer.set_wrap(
                    if paragraph.style.east_asian_line_break || is_east_asian_text(&text) {
                        Wrap::WordOrGlyph
                    } else {
                        Wrap::Word
                    },
                );
            }
            buffer.set_hinting(Hinting::Disabled);
            let alignment_width = if centered_hanging_paragraph {
                inner_width
            } else {
                available_width
            };
            buffer.set_size(Some(alignment_width), Some(100_000.0));
            buffer.shape_until_scroll(font_system, false);
            buffers.push((buffer, Vec::new(), 0.0));
        }

        let mut bullet_buffer = paragraph.bullet.as_ref().map(|bullet| {
            let mut bullet_style = first_style.clone();
            bullet_style.color = bullet.color.clone();
            if let Some(font_family) = &bullet.font_family {
                bullet_style.font_family = font_family.clone();
                bullet_style.east_asian_font_family.clear();
            }
            let bullet_font_size = bullet.font_size.unwrap_or(font_size).max(1.0);
            bullet_style.font_size = bullet_font_size;
            let bullet_attrs = cosmic_attrs(&bullet_style, &bullet.char, scale, line_height);
            let attrs = AttrsList::new(&bullet_attrs);
            let bullet_line =
                BufferLine::new(&bullet.char, LineEnding::None, attrs, Shaping::Advanced);
            let mut bullet_buffer = Buffer::new_empty(Metrics::new(bullet_font_size, line_height));
            bullet_buffer.lines.push(bullet_line);
            bullet_buffer.set_hinting(Hinting::Disabled);
            bullet_buffer.set_size(Some(bullet_font_size * 4.0), Some(line_height * 2.0));
            bullet_buffer.set_wrap(Wrap::None);
            bullet_buffer.shape_until_scroll(font_system, false);
            bullet_buffer
        });

        let buffer_count = buffers.len();
        let mut paragraph_height = if vertical_text {
            vertical_line_height
        } else {
            line_height
        };
        for (chunk_index, (buffer, vertical_line_offsets, vertical_height)) in
            buffers.into_iter().enumerate()
        {
            let chunk_height = if vertical_text {
                vertical_height.max(vertical_line_height)
            } else {
                buffer
                    .layout_runs()
                    .map(|run| run.line_top + run.line_height)
                    .fold(line_height, f32::max)
            };
            paragraph_height = paragraph_height.max(chunk_height);
            let column_index = vertical_column + chunk_index;
            let x = if vertical_text {
                let column_width = font_size.max(1.0);
                let right_edge = margin_left + inner_width - column_width;
                right_edge - column_index as f32 * column_width
            } else if centered_hanging_paragraph {
                margin_left
            } else {
                margin_left + paragraph_left
            };
            let top = if vertical_text {
                let available_height = (txt.rect.h - margin_top - margin_bottom).max(0.0);
                let anchor_offset =
                    match txt.body.as_ref().map(|body| body.vertical_anchor.as_str()) {
                        Some("middle") => ((available_height - vertical_height) / 2.0).max(0.0),
                        Some("bottom") => (available_height - vertical_height).max(0.0),
                        _ => 0.0,
                    };
                margin_top + anchor_offset
            } else {
                cursor_y
            };
            layouts.push(CosmicParagraph {
                buffer,
                bullet_buffer: if chunk_index == 0 {
                    bullet_buffer.take()
                } else {
                    None
                },
                x,
                column_width: if vertical_text {
                    font_size.max(1.0)
                } else {
                    0.0
                },
                vertical_line_offsets: if vertical_text {
                    vertical_line_offsets
                } else {
                    Vec::new()
                },
                vertical_column_height: if vertical_text { vertical_height } else { 0.0 },
                first_line_offset,
                subsequent_line_offset,
                bullet_x: margin_left + paragraph_left + paragraph.style.indent,
                top,
                hanging_punctuation: paragraph.style.hanging_punctuation,
                font_alignment: paragraph.style.font_alignment.clone(),
                font_size,
                metric_line_height,
                line_height,
                font_metrics,
            });
        }
        if vertical_text {
            vertical_column += buffer_count;
            vertical_content_height = vertical_content_height.max(paragraph_height);
        } else {
            cursor_y += paragraph_height + paragraph.style.space_after * scale;
        }
    }

    if vertical_text {
        // The current x calculation naturally builds columns from the right
        // edge, matching eaVert order. Center the complete group afterward so
        // narrow boxes do not make every column appear right-aligned or clamp
        // multiple columns to the same left edge.
        if let (Some(min_x), Some(max_right)) = (
            layouts.iter().map(|layout| layout.x).reduce(f32::min),
            layouts
                .iter()
                .map(|layout| layout.x + layout.column_width)
                .reduce(f32::max),
        ) {
            let group_center = (min_x + max_right) * 0.5;
            let box_center = margin_left + inner_width * 0.5;
            let offset = box_center - group_center;
            for layout in &mut layouts {
                layout.x += offset;
            }
        }
        (
            layouts,
            margin_top + vertical_content_height + margin_bottom,
        )
    } else {
        (layouts, cursor_y + margin_bottom)
    }
}
