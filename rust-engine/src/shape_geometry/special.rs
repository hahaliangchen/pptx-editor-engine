use super::helpers::{polygon, rounded_rect};
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
        "arc" => {
            ctx.move_to(x + w * 0.1, y + h * 0.82);
            let _ = ctx.arc(x + w * 0.5, y + h * 0.5, w.min(h) * 0.42, 0.4, 2.7);
        }
        "chord" => {
            ctx.move_to(x + w * 0.1, y + h * 0.82);
            let _ = ctx.arc(x + w * 0.5, y + h * 0.5, w.min(h) * 0.42, 0.4, 2.7);
            ctx.close_path();
        }
        "teardrop" => {
            ctx.move_to(x + w / 2.0, y);
            ctx.bezier_curve_to(
                x + w * 1.05,
                y + h * 0.55,
                x + w * 0.82,
                y + h,
                x + w / 2.0,
                y + h,
            );
            ctx.bezier_curve_to(
                x + w * 0.18,
                y + h,
                x - w * 0.05,
                y + h * 0.55,
                x + w / 2.0,
                y,
            );
            ctx.close_path();
        }
        "homePlate" => polygon(
            ctx,
            &[
                (x, y),
                (x + w * 0.7, y),
                (x + w, y + h / 2.0),
                (x + w * 0.7, y + h),
                (x, y + h),
            ],
        ),
        "chevron" => polygon(
            ctx,
            &[
                (x, y),
                (x + w * 0.62, y),
                (x + w, y + h / 2.0),
                (x + w * 0.62, y + h),
                (x, y + h),
                (x + w * 0.38, y + h / 2.0),
            ],
        ),
        "pieWedge" | "pie" => {
            let _ = ctx.move_to(x + w / 2.0, y + h / 2.0);
            let _ = ctx.arc(x + w / 2.0, y + h / 2.0, w.min(h) / 2.0, -0.6, 1.8);
            ctx.close_path();
        }
        "blockArc" => {
            let _ = ctx.arc(x + w / 2.0, y + h / 2.0, w.min(h) / 2.0, -0.8, 1.8);
            let _ = ctx.arc(x + w / 2.0, y + h / 2.0, w.min(h) * 0.28, 1.8, -0.8);
            ctx.close_path();
        }
        "donut" => {
            let _ = ctx.ellipse(
                x + w / 2.0,
                y + h / 2.0,
                w / 2.0,
                h / 2.0,
                0.0,
                0.0,
                std::f64::consts::TAU,
            );
            let _ = ctx.ellipse(
                x + w / 2.0,
                y + h / 2.0,
                w * 0.25,
                h * 0.25,
                0.0,
                0.0,
                std::f64::consts::TAU,
            );
        }
        "noSmoking" => {
            let _ = ctx.ellipse(
                x + w / 2.0,
                y + h / 2.0,
                w / 2.0,
                h / 2.0,
                0.0,
                0.0,
                std::f64::consts::TAU,
            );
            ctx.move_to(x + w * 0.18, y + h * 0.18);
            ctx.line_to(x + w * 0.82, y + h * 0.82);
        }
        "cloud" | "cloudCallout" => {
            ctx.move_to(x + w * 0.18, y + h * 0.7);
            ctx.bezier_curve_to(
                x - w * 0.05,
                y + h * 0.45,
                x + w * 0.1,
                y + h * 0.18,
                x + w * 0.32,
                y + h * 0.28,
            );
            ctx.bezier_curve_to(
                x + w * 0.38,
                y - h * 0.02,
                x + w * 0.72,
                y + h * 0.05,
                x + w * 0.74,
                y + h * 0.28,
            );
            ctx.bezier_curve_to(
                x + w * 1.04,
                y + h * 0.2,
                x + w * 1.04,
                y + h * 0.72,
                x + w * 0.78,
                y + h * 0.7,
            );
            ctx.bezier_curve_to(
                x + w * 0.65,
                y + h,
                x + w * 0.3,
                y + h * 0.95,
                x + w * 0.18,
                y + h * 0.7,
            );
            ctx.close_path();
        }
        "ribbon" | "ribbon2" | "ellipseRibbon" | "ellipseRibbon2" | "leftRightRibbon" => {
            polygon(
                ctx,
                &[
                    (x, y + h * 0.22),
                    (x + w * 0.18, y),
                    (x + w * 0.82, y),
                    (x + w, y + h * 0.22),
                    (x + w * 0.82, y + h),
                    (x + w * 0.18, y + h),
                ],
            );
        }
        "verticalScroll" | "horizontalScroll" => rounded_rect(ctx, x, y, w, h, 0.18),
        "wave" | "doubleWave" => {
            ctx.move_to(x, y + h * 0.55);
            ctx.bezier_curve_to(
                x + w * 0.2,
                y + h * 0.05,
                x + w * 0.3,
                y + h * 0.95,
                x + w * 0.5,
                y + h * 0.55,
            );
            ctx.bezier_curve_to(
                x + w * 0.7,
                y + h * 0.05,
                x + w * 0.8,
                y + h * 0.95,
                x + w,
                y + h * 0.55,
            );
            ctx.line_to(x + w, y + h);
            ctx.line_to(x, y + h);
            ctx.close_path();
        }
        _ => return false,
    }
    true
}
