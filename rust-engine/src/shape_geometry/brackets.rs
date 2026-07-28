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
    let left = kind == "leftBracket" || kind == "leftBrace";
    let right = kind == "rightBracket" || kind == "rightBrace";
    let pair = kind == "bracketPair" || kind == "bracePair";
    if !(left || right || pair) {
        return false;
    }
    let brace = kind.contains("Brace");
    let draw_one = |ctx: &CanvasRenderingContext2d, x: f64, flip: bool| {
        let mid = y + h / 2.0;
        if brace {
            let side = if flip { -1.0 } else { 1.0 };
            ctx.move_to(x + side * w * 0.18, y);
            ctx.bezier_curve_to(x, y + h * 0.18, x, y + h * 0.32, x + side * w * 0.12, mid);
            ctx.bezier_curve_to(x, y + h * 0.68, x, y + h * 0.82, x + side * w * 0.18, y + h);
        } else {
            ctx.move_to(x + if flip { -w * 0.15 } else { w * 0.15 }, y);
            ctx.line_to(x, y);
            ctx.line_to(x, y + h);
            ctx.line_to(x + if flip { -w * 0.15 } else { w * 0.15 }, y + h);
        }
    };
    if pair {
        draw_one(ctx, x + w * 0.18, false);
        draw_one(ctx, x + w * 0.82, true);
    } else if left {
        draw_one(ctx, x + w * 0.18, false);
    } else if right {
        draw_one(ctx, x + w * 0.82, true);
    }
    true
}
