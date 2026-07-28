use super::helpers::{adjustment, polygon, rounded_rect, snipped_rect};
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
    match shp.shape_type.as_str() {
        "rect" => ctx.rect(x, y, w, h),
        "roundRect" => rounded_rect(
            ctx,
            x,
            y,
            w,
            h,
            shp.corner_radius.unwrap_or(1.0 / 6.0) as f64,
        ),
        "ellipse" | "oval" | "flowChartConnector" => {
            let _ = ctx.ellipse(
                x + w / 2.0,
                y + h / 2.0,
                w / 2.0,
                h / 2.0,
                0.0,
                0.0,
                std::f64::consts::TAU,
            );
        }
        "line" | "lineInv" => {
            if shp.shape_type == "line" {
                ctx.move_to(x, y);
                ctx.line_to(x + w, y + h);
            } else {
                ctx.move_to(x, y + h);
                ctx.line_to(x + w, y);
            }
        }
        "triangle" => polygon(ctx, &[(x + w / 2.0, y), (x + w, y + h), (x, y + h)]),
        "rtTriangle" => polygon(ctx, &[(x, y + h), (x, y), (x + w, y + h)]),
        "diamond" => polygon(
            ctx,
            &[
                (x + w / 2.0, y),
                (x + w, y + h / 2.0),
                (x + w / 2.0, y + h),
                (x, y + h / 2.0),
            ],
        ),
        "parallelogram" => {
            let slant = w * adjustment(shp, "adj", 0.25);
            polygon(
                ctx,
                &[
                    (x + slant, y),
                    (x + w, y),
                    (x + w - slant, y + h),
                    (x, y + h),
                ],
            );
        }
        "trapezoid" | "nonIsoscelesTrapezoid" => {
            let inset = w * adjustment(shp, "adj", 0.25);
            polygon(
                ctx,
                &[
                    (x + inset, y),
                    (x + w - inset, y),
                    (x + w, y + h),
                    (x, y + h),
                ],
            );
        }
        "round1Rect" | "round2SameRect" | "round2DiagRect" => rounded_rect(ctx, x, y, w, h, 0.16),
        "snipRoundRect" | "snip1Rect" | "snip2SameRect" | "snip2DiagRect" => {
            snipped_rect(ctx, x, y, w, h, 0.16)
        }
        "plaque" => rounded_rect(ctx, x, y, w, h, 0.28),
        _ => return false,
    }
    true
}

pub fn contains(shp: &ShapeElement, px: f32, py: f32, w: f32, h: f32) -> Option<bool> {
    match shp.shape_type.as_str() {
        "ellipse" | "oval" | "flowChartConnector" => {
            let dx = (px - w / 2.0) / (w / 2.0).max(0.001);
            let dy = (py - h / 2.0) / (h / 2.0).max(0.001);
            Some(dx * dx + dy * dy <= 1.0)
        }
        "triangle" => Some((px - w / 2.0).abs() <= (py / h.max(0.001)) * w / 2.0),
        "diamond" => Some(
            (px - w / 2.0).abs() / (w / 2.0).max(0.001)
                + (py - h / 2.0).abs() / (h / 2.0).max(0.001)
                <= 1.0,
        ),
        "line" | "lineInv" => Some(true),
        "rect" | "roundRect" | "round1Rect" | "round2SameRect" | "round2DiagRect" | "plaque"
        | "snipRoundRect" | "snip1Rect" | "snip2SameRect" | "snip2DiagRect" => Some(true),
        _ => None,
    }
}
