use super::helpers::{polygon, rounded_rect, star};
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
        "mathPlus" | "plus" => {
            let arm = w.min(h) / 3.0;
            let l = x + (w - arm) / 2.0;
            let t = y + (h - arm) / 2.0;
            polygon(
                ctx,
                &[
                    (l, y),
                    (l + arm, y),
                    (l + arm, t),
                    (x + w, t),
                    (x + w, t + arm),
                    (l + arm, t + arm),
                    (l + arm, y + h),
                    (l, y + h),
                    (l, t + arm),
                    (x, t + arm),
                    (x, t),
                    (l, t),
                ],
            );
        }
        "mathMinus" => {
            ctx.rect(x, y + h * 0.4, w, h * 0.2);
        }
        "mathMultiply" => {
            ctx.move_to(x + w * 0.18, y);
            ctx.line_to(x + w * 0.5, y + h * 0.32);
            ctx.line_to(x + w * 0.82, y);
            ctx.line_to(x + w, y + h * 0.18);
            ctx.line_to(x + w * 0.68, y + h * 0.5);
            ctx.line_to(x + w, y + h * 0.82);
            ctx.line_to(x + w * 0.82, y + h);
            ctx.line_to(x + w * 0.5, y + h * 0.68);
            ctx.line_to(x + w * 0.18, y + h);
            ctx.line_to(x, y + h * 0.82);
            ctx.line_to(x + w * 0.32, y + h * 0.5);
            ctx.line_to(x, y + h * 0.18);
            ctx.close_path();
        }
        "mathDivide" => {
            let _ = ctx.ellipse(
                x + w / 2.0,
                y + h * 0.18,
                w * 0.1,
                h * 0.1,
                0.0,
                0.0,
                std::f64::consts::TAU,
            );
            ctx.rect(x, y + h * 0.4, w, h * 0.2);
            let _ = ctx.ellipse(
                x + w / 2.0,
                y + h * 0.82,
                w * 0.1,
                h * 0.1,
                0.0,
                0.0,
                std::f64::consts::TAU,
            );
        }
        "mathEqual" => {
            ctx.rect(x, y + h * 0.28, w, h * 0.16);
            ctx.rect(x, y + h * 0.56, w, h * 0.16);
        }
        "mathNotEqual" => {
            ctx.rect(x, y + h * 0.28, w, h * 0.16);
            ctx.rect(x, y + h * 0.56, w, h * 0.16);
            ctx.move_to(x + w * 0.72, y);
            ctx.line_to(x + w * 0.28, y + h);
        }
        "gear6" => star(ctx, x, y, w, h, 12, 0.76),
        "gear9" => star(ctx, x, y, w, h, 18, 0.8),
        "funnel" => polygon(
            ctx,
            &[
                (x, y),
                (x + w, y),
                (x + w * 0.62, y + h * 0.58),
                (x + w * 0.62, y + h),
                (x + w * 0.38, y + h),
                (x + w * 0.38, y + h * 0.58),
            ],
        ),
        "cube" => polygon(
            ctx,
            &[
                (x + w * 0.5, y),
                (x + w, y + h * 0.25),
                (x + w, y + h * 0.75),
                (x + w * 0.5, y + h),
                (x, y + h * 0.75),
                (x, y + h * 0.25),
            ],
        ),
        "can" => {
            let _ = ctx.ellipse(
                x + w / 2.0,
                y + h * 0.12,
                w / 2.0,
                h * 0.12,
                0.0,
                0.0,
                std::f64::consts::TAU,
            );
            ctx.rect(x, y + h * 0.12, w, h * 0.76);
            let _ = ctx.ellipse(
                x + w / 2.0,
                y + h * 0.88,
                w / 2.0,
                h * 0.12,
                0.0,
                0.0,
                std::f64::consts::TAU,
            );
        }
        "heart" => {
            ctx.move_to(x + w / 2.0, y + h);
            ctx.bezier_curve_to(x, y + h * 0.62, x, y + h * 0.12, x + w * 0.25, y + h * 0.12);
            ctx.bezier_curve_to(
                x + w * 0.42,
                y + h * 0.12,
                x + w * 0.5,
                y + h * 0.3,
                x + w / 2.0,
                y + h * 0.34,
            );
            ctx.bezier_curve_to(
                x + w * 0.5,
                y + h * 0.3,
                x + w * 0.58,
                y + h * 0.12,
                x + w * 0.75,
                y + h * 0.12,
            );
            ctx.bezier_curve_to(x + w, y + h * 0.12, x + w, y + h * 0.62, x + w / 2.0, y + h);
            ctx.close_path();
        }
        "lightningBolt" => polygon(
            ctx,
            &[
                (x + w * 0.58, y),
                (x + w * 0.1, y + h * 0.56),
                (x + w * 0.46, y + h * 0.56),
                (x + w * 0.34, y + h),
                (x + w * 0.9, y + h * 0.36),
                (x + w * 0.56, y + h * 0.36),
            ],
        ),
        "sun" => star(ctx, x, y, w, h, 16, 0.62),
        "moon" => {
            ctx.move_to(x + w * 0.72, y + h * 0.08);
            ctx.bezier_curve_to(
                x + w * 0.18,
                y,
                x + w * 0.12,
                y + h * 0.82,
                x + w * 0.78,
                y + h * 0.92,
            );
            ctx.bezier_curve_to(
                x + w * 0.42,
                y + h * 0.62,
                x + w * 0.42,
                y + h * 0.3,
                x + w * 0.72,
                y + h * 0.08,
            );
            ctx.close_path();
        }
        "smileyFace" => {
            let _ = ctx.ellipse(
                x + w / 2.0,
                y + h / 2.0,
                w / 2.0,
                h / 2.0,
                0.0,
                0.0,
                std::f64::consts::TAU,
            );
            let _ = ctx.arc(
                x + w * 0.35,
                y + h * 0.36,
                w * 0.05,
                0.0,
                std::f64::consts::TAU,
            );
            let _ = ctx.arc(
                x + w * 0.65,
                y + h * 0.36,
                w * 0.05,
                0.0,
                std::f64::consts::TAU,
            );
            ctx.move_to(x + w * 0.28, y + h * 0.62);
            ctx.bezier_curve_to(
                x + w * 0.4,
                y + h * 0.78,
                x + w * 0.6,
                y + h * 0.78,
                x + w * 0.72,
                y + h * 0.62,
            );
        }
        "foldedCorner" => polygon(
            ctx,
            &[
                (x, y),
                (x + w * 0.72, y),
                (x + w, y + h * 0.28),
                (x + w, y + h),
                (x, y + h),
            ],
        ),
        "bevel" | "frame" | "halfFrame" | "corner" | "diagStripe" => {
            rounded_rect(ctx, x, y, w, h, 0.08)
        }
        "cornerTabs" | "squareTabs" | "plaqueTabs" | "chartX" | "chartStar" | "chartPlus" => {
            rounded_rect(ctx, x, y, w, h, 0.1)
        }
        "actionButtonBlank"
        | "actionButtonHome"
        | "actionButtonHelp"
        | "actionButtonInformation"
        | "actionButtonForwardNext"
        | "actionButtonBackPrevious"
        | "actionButtonEnd"
        | "actionButtonBeginning"
        | "actionButtonReturn"
        | "actionButtonDocument"
        | "actionButtonSound"
        | "actionButtonMovie" => rounded_rect(ctx, x, y, w, h, 0.22),
        _ => return false,
    }
    true
}
