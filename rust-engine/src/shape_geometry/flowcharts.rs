use super::helpers::{polygon, regular_polygon, rounded_rect};
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
        "flowChartProcess" => ctx.rect(x, y, w, h),
        "flowChartDecision" | "flowChartSort" => polygon(
            ctx,
            &[
                (x + w / 2.0, y),
                (x + w, y + h / 2.0),
                (x + w / 2.0, y + h),
                (x, y + h / 2.0),
            ],
        ),
        "flowChartInputOutput" | "flowChartManualOperation" => {
            let slant = w * 0.16;
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
        "flowChartPredefinedProcess" => {
            ctx.rect(x, y, w, h);
            ctx.move_to(x + w * 0.14, y);
            ctx.line_to(x + w * 0.14, y + h);
            ctx.move_to(x + w * 0.86, y);
            ctx.line_to(x + w * 0.86, y + h);
        }
        "flowChartInternalStorage" => {
            ctx.rect(x, y, w, h);
            ctx.move_to(x + w * 0.2, y);
            ctx.line_to(x + w * 0.2, y + h);
            ctx.move_to(x, y + h * 0.2);
            ctx.line_to(x + w, y + h * 0.2);
        }
        "flowChartDocument" | "flowChartMultidocument" => {
            ctx.move_to(x, y);
            ctx.line_to(x + w, y);
            ctx.line_to(x + w, y + h * 0.82);
            ctx.bezier_curve_to(
                x + w * 0.72,
                y + h * 1.05,
                x + w * 0.28,
                y + h * 0.65,
                x,
                y + h * 0.9,
            );
            ctx.close_path();
        }
        "flowChartTerminator" | "flowChartAlternateProcess" => rounded_rect(ctx, x, y, w, h, 0.45),
        "flowChartPreparation" => {
            let c = w.min(h) * 0.18;
            polygon(
                ctx,
                &[
                    (x + c, y),
                    (x + w - c, y),
                    (x + w, y + h / 2.0),
                    (x + w - c, y + h),
                    (x + c, y + h),
                    (x, y + h / 2.0),
                ],
            );
        }
        "flowChartManualInput" => polygon(
            ctx,
            &[(x, y + h * 0.25), (x + w, y), (x + w, y + h), (x, y + h)],
        ),
        "flowChartConnector" => {
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
        "flowChartPunchedCard" => polygon(
            ctx,
            &[
                (x, y + h * 0.18),
                (x + w * 0.18, y),
                (x + w, y),
                (x + w, y + h),
                (x, y + h),
            ],
        ),
        "flowChartPunchedTape" => polygon(
            ctx,
            &[
                (x, y + h * 0.16),
                (x + w * 0.18, y),
                (x + w, y + h * 0.12),
                (x + w, y + h * 0.84),
                (x + w * 0.82, y + h),
                (x, y + h * 0.88),
            ],
        ),
        "flowChartSummingJunction" | "flowChartOr" => {
            let _ = ctx.ellipse(
                x + w / 2.0,
                y + h / 2.0,
                w / 2.0,
                h / 2.0,
                0.0,
                0.0,
                std::f64::consts::TAU,
            );
            ctx.move_to(x + w * 0.25, y + h / 2.0);
            ctx.line_to(x + w * 0.75, y + h / 2.0);
            ctx.move_to(x + w / 2.0, y + h * 0.25);
            ctx.line_to(x + w / 2.0, y + h * 0.75);
        }
        "flowChartCollate" => {
            regular_polygon(ctx, x, y, w, h, 6, 0.0);
            ctx.move_to(x + w * 0.28, y + h * 0.2);
            ctx.line_to(x + w * 0.72, y + h * 0.8);
        }
        "flowChartExtract" => polygon(
            ctx,
            &[
                (x + w / 2.0, y),
                (x + w, y + h / 2.0),
                (x + w / 2.0, y + h),
                (x, y + h / 2.0),
            ],
        ),
        "flowChartMerge" => polygon(ctx, &[(x, y), (x + w, y), (x + w / 2.0, y + h)]),
        "flowChartOfflineStorage" | "flowChartOnlineStorage" => {
            ctx.move_to(x, y);
            ctx.line_to(x + w, y);
            ctx.line_to(x + w, y + h * 0.82);
            ctx.bezier_curve_to(x + w * 0.75, y + h, x + w * 0.25, y + h, x, y + h * 0.82);
            ctx.close_path();
        }
        "flowChartMagneticTape" | "flowChartMagneticDisk" | "flowChartMagneticDrum" => {
            let _ = ctx.ellipse(
                x + w / 2.0,
                y + h * 0.14,
                w / 2.0,
                h * 0.14,
                0.0,
                0.0,
                std::f64::consts::TAU,
            );
            ctx.rect(x, y + h * 0.14, w, h * 0.68);
            let _ = ctx.ellipse(
                x + w / 2.0,
                y + h * 0.82,
                w / 2.0,
                h * 0.14,
                0.0,
                0.0,
                std::f64::consts::TAU,
            );
        }
        "flowChartDisplay" => {
            ctx.move_to(x + w * 0.12, y);
            ctx.line_to(x + w * 0.88, y);
            ctx.line_to(x + w, y + h * 0.5);
            ctx.line_to(x + w * 0.88, y + h);
            ctx.line_to(x + w * 0.12, y + h);
            ctx.line_to(x, y + h * 0.5);
            ctx.close_path();
        }
        "flowChartDelay" => rounded_rect(ctx, x, y, w, h, 0.4),
        "flowChartOffpageConnector" => polygon(
            ctx,
            &[
                (x, y),
                (x + w, y),
                (x + w, y + h * 0.72),
                (x + w / 2.0, y + h),
                (x, y + h * 0.72),
            ],
        ),
        _ => return false,
    }
    true
}

pub fn contains(shp: &ShapeElement, _px: f32, _py: f32, _w: f32, _h: f32) -> Option<bool> {
    if shp.shape_type.starts_with("flowChart") {
        Some(true)
    } else {
        None
    }
}
