use super::helpers::line;
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
    let kind = shp.shape_type.as_str();
    if kind == "straightConnector1" {
        line(ctx, x, y, x + w, y + h);
    } else if kind.starts_with("bentConnector") {
        ctx.move_to(x, y);
        ctx.line_to(x + w / 2.0, y);
        ctx.line_to(x + w / 2.0, y + h);
        ctx.line_to(x + w, y + h);
    } else if kind.starts_with("curvedConnector") {
        ctx.move_to(x, y);
        ctx.bezier_curve_to(x + w * 0.35, y, x + w * 0.65, y + h, x + w, y + h);
    } else {
        return false;
    }
    true
}

pub fn contains(shp: &ShapeElement, _px: f32, _py: f32, _w: f32, _h: f32) -> Option<bool> {
    if shp.shape_type.contains("Connector") {
        Some(true)
    } else {
        None
    }
}
