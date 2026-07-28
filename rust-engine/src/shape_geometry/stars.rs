use super::helpers::star;
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
    let points = match shp.shape_type.as_str() {
        "star4" => 4,
        "star5" => 5,
        "star6" => 6,
        "star7" => 7,
        "star8" => 8,
        "star10" => 10,
        "star12" => 12,
        "star16" => 16,
        "star24" => 24,
        "star32" => 32,
        "irregularSeal1" => 16,
        "irregularSeal2" => 24,
        _ => return false,
    };
    let inner = if shp.shape_type.starts_with("irregularSeal") {
        0.82
    } else {
        0.44
    };
    star(ctx, x, y, w, h, points, inner);
    true
}

pub fn contains(shp: &ShapeElement, _px: f32, _py: f32, _w: f32, _h: f32) -> Option<bool> {
    if shp.shape_type.starts_with("star") || shp.shape_type.starts_with("irregularSeal") {
        Some(true)
    } else {
        None
    }
}
