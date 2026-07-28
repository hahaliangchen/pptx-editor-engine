use super::helpers::rounded_rect;
use crate::ast::ShapeElement;
use web_sys::CanvasRenderingContext2d;

fn is_callout(kind: &str) -> bool {
    kind.starts_with("callout")
        || kind.starts_with("accentCallout")
        || kind.starts_with("borderCallout")
        || kind.starts_with("accentBorderCallout")
        || kind.starts_with("wedge")
        || kind == "cloudCallout"
}

pub fn draw(
    ctx: &CanvasRenderingContext2d,
    shp: &ShapeElement,
    x: f64,
    y: f64,
    w: f64,
    h: f64,
) -> bool {
    let kind = shp.shape_type.as_str();
    if !is_callout(kind) {
        return false;
    }
    if kind == "cloudCallout" {
        ctx.move_to(x + w * 0.18, y + h * 0.65);
        ctx.bezier_curve_to(
            x - w * 0.05,
            y + h * 0.45,
            x + w * 0.08,
            y + h * 0.18,
            x + w * 0.3,
            y + h * 0.28,
        );
        ctx.bezier_curve_to(
            x + w * 0.35,
            y - h * 0.02,
            x + w * 0.7,
            y + h * 0.04,
            x + w * 0.72,
            y + h * 0.27,
        );
        ctx.bezier_curve_to(
            x + w * 1.05,
            y + h * 0.22,
            x + w * 1.05,
            y + h * 0.7,
            x + w * 0.78,
            y + h * 0.68,
        );
        ctx.bezier_curve_to(
            x + w * 0.75,
            y + h * 0.98,
            x + w * 0.42,
            y + h * 0.96,
            x + w * 0.38,
            y + h * 0.72,
        );
        ctx.close_path();
    } else {
        rounded_rect(ctx, x, y, w, h * 0.78, 0.12);
        ctx.move_to(x + w * 0.28, y + h * 0.78);
        ctx.line_to(x + w * 0.18, y + h);
        ctx.line_to(x + w * 0.42, y + h * 0.78);
        ctx.close_path();
    }
    true
}

pub fn contains(shp: &ShapeElement, _px: f32, _py: f32, _w: f32, _h: f32) -> Option<bool> {
    if is_callout(&shp.shape_type) {
        Some(true)
    } else {
        None
    }
}
