#![allow(dead_code)]
use serde::{Deserialize, Serialize};

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct Presentation {
    pub size: Size,
    pub slides: Vec<Slide>,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct Size {
    pub width: f32,
    pub height: f32,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct Slide {
    pub id: String,
    pub elements: Vec<Element>,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(tag = "type")]
pub enum Element {
    #[serde(rename = "text")]
    Text(TextElement),
    #[serde(rename = "shape")]
    Shape(ShapeElement),
    #[serde(rename = "image")]
    Image(ImageElement),
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct Rect {
    pub x: f32,
    pub y: f32,
    pub w: f32,
    pub h: f32,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct TextElement {
    pub id: String,
    pub rect: Rect,
    pub content: String,
    pub style: TextStyle,
    #[serde(default)]
    pub paragraphs: Vec<TextParagraph>,
    #[serde(default)]
    pub body: Option<TextBodyProperties>,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
pub struct TextStyle {
    #[serde(rename = "fontSize")]
    pub font_size: f32,
    pub color: String,
    pub bold: bool,
    #[serde(default = "default_align")]
    pub align: String, // "left" | "center" | "right"
    #[serde(rename = "fontFamily", default = "default_font_family")]
    pub font_family: String,
    #[serde(rename = "eastAsianFontFamily", default)]
    pub east_asian_font_family: String,
    #[serde(default)]
    pub italic: bool,
    #[serde(rename = "letterSpacing", default)]
    pub letter_spacing: f32,
}

fn default_align() -> String {
    "left".to_string()
}

fn default_font_family() -> String {
    "sans-serif".to_string()
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct TextRun {
    pub content: String,
    pub style: TextStyle,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct ParagraphStyle {
    pub align: String,
    pub level: u8,
    #[serde(rename = "marginLeft")]
    pub margin_left: f32,
    pub indent: f32,
    #[serde(rename = "eastAsianLineBreak", default)]
    pub east_asian_line_break: bool,
    #[serde(rename = "hangingPunctuation", default)]
    pub hanging_punctuation: bool,
    #[serde(rename = "fontAlignment", default = "default_font_alignment")]
    pub font_alignment: String,
    #[serde(rename = "lineSpacing", default)]
    pub line_spacing: Option<LineSpacing>,
    #[serde(rename = "spaceBefore", default)]
    pub space_before: f32,
    #[serde(rename = "spaceAfter", default)]
    pub space_after: f32,
}

fn default_font_alignment() -> String {
    "auto".to_string()
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct LineSpacing {
    pub unit: String,
    pub value: f32,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct Bullet {
    pub char: String,
    pub color: String,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct TextParagraph {
    pub runs: Vec<TextRun>,
    pub style: ParagraphStyle,
    #[serde(default)]
    pub bullet: Option<Bullet>,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct TextBodyProperties {
    #[serde(rename = "marginLeft")]
    pub margin_left: f32,
    #[serde(rename = "marginRight")]
    pub margin_right: f32,
    #[serde(rename = "marginTop")]
    pub margin_top: f32,
    #[serde(rename = "marginBottom")]
    pub margin_bottom: f32,
    #[serde(rename = "verticalAnchor")]
    pub vertical_anchor: String,
    #[serde(rename = "autoFit")]
    pub auto_fit: String,
    #[serde(rename = "fontScale", default = "default_font_scale")]
    pub font_scale: f32,
}

fn default_font_scale() -> f32 {
    1.0
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct Border {
    pub color: String,
    pub width: f32,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(tag = "type")]
pub enum FillStyle {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "solid")]
    Solid { color: String },
    #[serde(rename = "gradient")]
    Gradient {
        kind: String,
        stops: Vec<GradientStop>,
        #[serde(default)]
        angle: Option<f32>,
        #[serde(rename = "rotateWithShape", default)]
        rotate_with_shape: Option<bool>,
    },
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct GradientStop {
    pub position: f32,
    pub color: String,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct LineStyle {
    pub fill: FillStyle,
    pub width: f32,
    #[serde(default)]
    pub dash: Option<String>,
    #[serde(default)]
    pub cap: Option<String>,
    #[serde(default)]
    pub join: Option<String>,
    #[serde(default)]
    pub compound: Option<String>,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct ShadowStyle {
    pub color: String,
    pub opacity: f32,
    pub blur: f32,
    pub distance: f32,
    pub direction: f32,
    #[serde(rename = "scaleX", default = "default_shadow_scale")]
    pub scale_x: f32,
    #[serde(rename = "scaleY", default = "default_shadow_scale")]
    pub scale_y: f32,
    #[serde(default)]
    pub alignment: Option<String>,
}

fn default_shadow_scale() -> f32 {
    1.0
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct GlowStyle {
    pub color: String,
    pub opacity: f32,
    pub radius: f32,
}

#[derive(Serialize, Deserialize, Debug, Clone, Default)]
pub struct EffectStyle {
    #[serde(rename = "outerShadow", default)]
    pub outer_shadow: Option<ShadowStyle>,
    #[serde(default)]
    pub glow: Option<GlowStyle>,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct ComputedShapeStyle {
    pub fill: FillStyle,
    #[serde(default)]
    pub line: Option<LineStyle>,
    #[serde(default)]
    pub effects: Option<EffectStyle>,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct ShapeElement {
    pub id: String,
    pub rect: Rect,
    #[serde(rename = "shapeType")]
    pub shape_type: String, // "rect" | "ellipse" | "triangle"
    pub fill: String,
    #[serde(default)]
    pub border: Option<Border>,
    #[serde(rename = "computedStyle", default)]
    pub computed_style: Option<ComputedShapeStyle>,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct ImageElement {
    pub id: String,
    pub rect: Rect,
    pub url: String, // Blob URL or base64 data URI
    #[serde(default)]
    pub crop: Option<ImageCrop>,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct ImageCrop {
    pub left: f32,
    pub top: f32,
    pub right: f32,
    pub bottom: f32,
}

#[cfg(test)]
mod tests {
    use super::{Element, Slide};

    #[test]
    fn legacy_text_ast_keeps_working() {
        let json = r##"{
            "id":"slide_1",
            "elements":[{
                "type":"text",
                "id":"text_1",
                "rect":{"x":0,"y":0,"w":100,"h":30},
                "content":"Title",
                "style":{"fontSize":24,"color":"#000000","bold":true,"align":"center"}
            }]
        }"##;

        let slide: Slide = serde_json::from_str(json).expect("legacy AST should deserialize");
        let Element::Text(text) = &slide.elements[0] else {
            panic!("expected text element");
        };
        assert_eq!(text.style.font_family, "sans-serif");
        assert!(text.paragraphs.is_empty());
        assert!(text.body.is_none());
    }

    #[test]
    fn computed_text_ast_deserializes() {
        let json = r##"{
            "id":"slide_1",
            "elements":[{
                "type":"text",
                "id":"text_1",
                "rect":{"x":8,"y":4,"w":84,"h":30},
                "content":"Title",
                "style":{"fontSize":24,"color":"#000000","bold":true,"align":"center","fontFamily":"Aptos","italic":false},
                "paragraphs":[{
                    "style":{"align":"center","level":0,"marginLeft":0,"indent":0,"lineSpacing":{"unit":"percent","value":1.2},"spaceBefore":2,"spaceAfter":3},
                    "runs":[{"content":"Title","style":{"fontSize":24,"color":"#000000","bold":true,"align":"center","fontFamily":"Aptos"}}]
                }],
                "body":{"marginLeft":8,"marginRight":8,"marginTop":4,"marginBottom":4,"verticalAnchor":"top","autoFit":"none"}
            }]
        }"##;

        let slide: Slide = serde_json::from_str(json).expect("computed AST should deserialize");
        let Element::Text(text) = &slide.elements[0] else {
            panic!("expected text element");
        };
        assert_eq!(text.style.font_family, "Aptos");
        assert_eq!(text.paragraphs[0].runs[0].content, "Title");
        assert_eq!(
            text.paragraphs[0]
                .style
                .line_spacing
                .as_ref()
                .unwrap()
                .value,
            1.2
        );
        assert_eq!(text.paragraphs[0].style.space_after, 3.0);
        assert!(!text.paragraphs[0].style.east_asian_line_break);
        assert!(!text.paragraphs[0].style.hanging_punctuation);
        assert_eq!(text.paragraphs[0].style.font_alignment, "auto");
        assert_eq!(text.body.as_ref().unwrap().margin_left, 8.0);
    }

    #[test]
    fn computed_shape_style_deserializes() {
        let json = r##"{
            "id":"slide_1",
            "elements":[{
                "type":"shape",
                "id":"shape_1",
                "rect":{"x":10,"y":20,"w":300,"h":80},
                "shapeType":"roundRect",
                "fill":"#ffffff",
                "computedStyle":{
                    "fill":{"type":"gradient","kind":"radial","stops":[
                        {"position":0,"color":"#ffffff"},
                        {"position":1,"color":"#ff0000"}
                    ]},
                    "line":{"fill":{"type":"solid","color":"#222222"},"width":2},
                    "effects":{"outerShadow":{
                        "color":"#000000","opacity":0.4,"blur":8,"distance":3,"direction":45
                    }}
                }
            }]
        }"##;

        let slide: Slide = serde_json::from_str(json).expect("shape style should deserialize");
        let Element::Shape(shape) = &slide.elements[0] else {
            panic!("expected shape element");
        };
        let style = shape.computed_style.as_ref().unwrap();
        let super::FillStyle::Gradient { stops, .. } = &style.fill else {
            panic!("expected gradient fill");
        };
        assert_eq!(stops.len(), 2);
        assert_eq!(style.line.as_ref().unwrap().width, 2.0);
        assert_eq!(
            style
                .effects
                .as_ref()
                .unwrap()
                .outer_shadow
                .as_ref()
                .unwrap()
                .opacity,
            0.4
        );
    }
}
