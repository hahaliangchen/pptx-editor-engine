use super::helpers::regular_polygon;
use crate::ast::ShapeElement;
use web_sys::CanvasRenderingContext2d;

pub fn draw(
    ctx: &CanvasRenderingContext2d,
    shp: &ShapeElement,
    x: f64,
    y: f64,
    w: f64,
    h: f64,
) -> bool {
    let sides = match shp.shape_type.as_str() {
        "pentagon" => 5,
        "hexagon" => 6,
        "heptagon" => 7,
        "octagon" => 8,
        "decagon" => 10,
        "dodecagon" => 12,
        _ => return false,
    };
    regular_polygon(ctx, x, y, w, h, sides, 0.0);
    true
}

pub fn contains(shp: &ShapeElement, _px: f32, _py: f32, _w: f32, _h: f32) -> Option<bool> {
    match shp.shape_type.as_str() {
        "pentagon" | "hexagon" | "heptagon" | "octagon" | "decagon" | "dodecagon" => Some(true),
        _ => None,
    }
}
