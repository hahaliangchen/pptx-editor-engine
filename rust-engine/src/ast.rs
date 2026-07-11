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
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct TextStyle {
    #[serde(rename = "fontSize")]
    pub font_size: f32,
    pub color: String,
    pub bold: bool,
    #[serde(default = "default_align")]
    pub align: String, // "left" | "center" | "right"
}

fn default_align() -> String {
    "left".to_string()
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct ShapeElement {
    pub id: String,
    pub rect: Rect,
    #[serde(rename = "shapeType")]
    pub shape_type: String, // "rect" | "ellipse" | "triangle"
    pub fill: String,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct ImageElement {
    pub id: String,
    pub rect: Rect,
    pub url: String, // Blob URL or base64 data URI
}
