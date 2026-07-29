use super::helpers::polygon;
use crate::ast::ShapeElement;
use web_sys::CanvasRenderingContext2d;

fn mapped_right_arrow(
    ctx: &CanvasRenderingContext2d,
    x: f64,
    y: f64,
    w: f64,
    h: f64,
    direction: &str,
    head: f64,
    shaft: f64,
) {
    let head = (w * head.clamp(0.15, 0.8)).min(w * 0.8);
    let shaft_h = (h * shaft.clamp(0.1, 1.0)).min(h);
    let top = (h - shaft_h) / 2.0;
    let bottom = top + shaft_h;
    let points = [
        (0.0, top),
        (1.0 - head / w, top),
        (1.0 - head / w, 0.0),
        (1.0, 0.5),
        (1.0 - head / w, 1.0),
        (1.0 - head / w, bottom),
        (0.0, bottom),
    ];
    let map = |u: f64, v: f64| -> (f64, f64) {
        let (u, v) = match direction {
            "left" => (1.0 - u, v),
            "down" => (v, u),
            "up" => (1.0 - v, 1.0 - u),
            _ => (u, v),
        };
        (x + u * w, y + v * h)
    };
    let mapped: Vec<_> = points.iter().map(|(u, v)| map(*u, *v)).collect();
    polygon(ctx, &mapped);
}

fn up_arrow(ctx: &CanvasRenderingContext2d, shp: &ShapeElement, x: f64, y: f64, w: f64, h: f64) {
    let short_side = w.min(h);
    let head_height =
        (short_side * shp.arrow_head_height.unwrap_or(0.5).clamp(0.15, 0.8) as f64).min(h * 0.8);
    let shaft_width = (w * shp.arrow_shaft_width.unwrap_or(0.5).clamp(0.1, 1.0) as f64).min(w);
    let center = x + w / 2.0;
    let shaft_left = center - shaft_width / 2.0;
    let shaft_right = center + shaft_width / 2.0;
    polygon(
        ctx,
        &[
            (center, y),
            (x + w, y + head_height),
            (shaft_right, y + head_height),
            (shaft_right, y + h),
            (shaft_left, y + h),
            (shaft_left, y + head_height),
            (x, y + head_height),
        ],
    );
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
    let direction = match kind {
        "rightArrow" | "stripedRightArrow" | "notchedRightArrow" | "rightArrowCallout" => "right",
        "leftArrow" | "leftArrowCallout" => "left",
        "upArrow" | "upArrowCallout" => "up",
        "downArrow" | "downArrowCallout" => "down",
        "leftRightArrow" | "leftRightArrowCallout" => "right",
        "upDownArrow" | "upDownArrowCallout" => "up",
        "leftUpArrow" | "leftRightUpArrow" | "quadArrow" | "quadArrowCallout" => "up",
        "bentUpArrow"
        | "bentArrow"
        | "uturnArrow"
        | "circularArrow"
        | "leftCircularArrow"
        | "leftRightCircularArrow"
        | "curvedRightArrow"
        | "curvedLeftArrow"
        | "curvedUpArrow"
        | "curvedDownArrow"
        | "swooshArrow" => "right",
        _ => return false,
    };
    if matches!(
        kind,
        "bentArrow"
            | "uturnArrow"
            | "circularArrow"
            | "leftCircularArrow"
            | "leftRightCircularArrow"
            | "swooshArrow"
    ) {
        ctx.move_to(x + w * 0.15, y + h * 0.75);
        ctx.bezier_curve_to(
            x + w * 0.05,
            y + h * 0.1,
            x + w * 0.75,
            y + h * 0.05,
            x + w * 0.78,
            y + h * 0.5,
        );
        ctx.line_to(x + w * 0.72, y + h * 0.38);
        ctx.line_to(x + w * 0.92, y + h * 0.52);
        ctx.line_to(x + w * 0.72, y + h * 0.66);
        ctx.close_path();
    } else if kind == "leftRightArrow" || kind == "leftRightArrowCallout" {
        mapped_right_arrow(ctx, x, y, w / 2.0, h, "left", 0.45, 0.42);
        mapped_right_arrow(ctx, x + w / 2.0, y, w / 2.0, h, "right", 0.45, 0.42);
    } else if kind == "upDownArrow" || kind == "upDownArrowCallout" {
        mapped_right_arrow(ctx, x, y, w, h / 2.0, "up", 0.45, 0.42);
        mapped_right_arrow(ctx, x, y + h / 2.0, w, h / 2.0, "down", 0.45, 0.42);
    } else if kind == "upArrow" {
        up_arrow(ctx, shp, x, y, w, h);
    } else {
        mapped_right_arrow(
            ctx,
            x,
            y,
            w,
            h,
            direction,
            shp.arrow_head_height.unwrap_or(0.5) as f64,
            shp.arrow_shaft_width.unwrap_or(0.5) as f64,
        );
    }
    true
}

pub fn contains(shp: &ShapeElement, _px: f32, _py: f32, _w: f32, _h: f32) -> Option<bool> {
    if shp.shape_type.contains("Arrow") || shp.shape_type.contains("arrow") {
        Some(true)
    } else {
        None
    }
}
