use std::f64::consts::PI;
use web_sys::CanvasRenderingContext2d;

pub fn adjustment(shp: &crate::ast::ShapeElement, name: &str, default: f32) -> f64 {
    shp.adjustments
        .get(name)
        .copied()
        .unwrap_or(default)
        .clamp(0.0, 1.0) as f64
}

pub fn polygon(ctx: &CanvasRenderingContext2d, points: &[(f64, f64)]) {
    if let Some((x, y)) = points.first() {
        ctx.move_to(*x, *y);
        for (x, y) in &points[1..] {
            ctx.line_to(*x, *y);
        }
        ctx.close_path();
    }
}

pub fn regular_polygon(
    ctx: &CanvasRenderingContext2d,
    x: f64,
    y: f64,
    w: f64,
    h: f64,
    sides: usize,
    rotation: f64,
) {
    let sides = sides.max(3);
    let cx = x + w / 2.0;
    let cy = y + h / 2.0;
    let rx = w / 2.0;
    let ry = h / 2.0;
    let points: Vec<_> = (0..sides)
        .map(|i| {
            let angle = rotation - PI / 2.0 + i as f64 * 2.0 * PI / sides as f64;
            (cx + rx * angle.cos(), cy + ry * angle.sin())
        })
        .collect();
    polygon(ctx, &points);
}

pub fn star(
    ctx: &CanvasRenderingContext2d,
    x: f64,
    y: f64,
    w: f64,
    h: f64,
    points: usize,
    inner_ratio: f64,
) {
    let points = points.max(3);
    let cx = x + w / 2.0;
    let cy = y + h / 2.0;
    let rx = w / 2.0;
    let ry = h / 2.0;
    let mut vertices = Vec::with_capacity(points * 2);
    for i in 0..points * 2 {
        let radius = if i % 2 == 0 { 1.0 } else { inner_ratio };
        let angle = -PI / 2.0 + i as f64 * PI / points as f64;
        vertices.push((
            cx + rx * radius * angle.cos(),
            cy + ry * radius * angle.sin(),
        ));
    }
    polygon(ctx, &vertices);
}

pub fn rounded_rect(ctx: &CanvasRenderingContext2d, x: f64, y: f64, w: f64, h: f64, ratio: f64) {
    let r = (w.min(h) * ratio.clamp(0.0, 0.5)).min(w.min(h) / 2.0);
    ctx.move_to(x + r, y);
    ctx.line_to(x + w - r, y);
    ctx.quadratic_curve_to(x + w, y, x + w, y + r);
    ctx.line_to(x + w, y + h - r);
    ctx.quadratic_curve_to(x + w, y + h, x + w - r, y + h);
    ctx.line_to(x + r, y + h);
    ctx.quadratic_curve_to(x, y + h, x, y + h - r);
    ctx.line_to(x, y + r);
    ctx.quadratic_curve_to(x, y, x + r, y);
    ctx.close_path();
}

pub fn snipped_rect(ctx: &CanvasRenderingContext2d, x: f64, y: f64, w: f64, h: f64, cut: f64) {
    let c = (w.min(h) * cut.clamp(0.0, 0.5)).min(w.min(h) / 2.0);
    polygon(
        ctx,
        &[
            (x + c, y),
            (x + w - c, y),
            (x + w, y + c),
            (x + w, y + h - c),
            (x + w - c, y + h),
            (x + c, y + h),
            (x, y + h - c),
            (x, y + c),
        ],
    );
}

pub fn line(ctx: &CanvasRenderingContext2d, x1: f64, y1: f64, x2: f64, y2: f64) {
    ctx.move_to(x1, y1);
    ctx.line_to(x2, y2);
}
