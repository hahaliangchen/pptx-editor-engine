use crate::ast::{FillStyle, Rect, ShapeElement};
use js_sys::Array;
use wasm_bindgen::JsValue;
use web_sys::CanvasRenderingContext2d;

pub fn set_shape_paint(
    ctx: &CanvasRenderingContext2d,
    fill: &FillStyle,
    rect: &Rect,
    stroke: bool,
) -> bool {
    match fill {
        FillStyle::None => false,
        FillStyle::Solid { color } => {
            if stroke {
                ctx.set_stroke_style_str(color);
            } else {
                ctx.set_fill_style_str(color);
            }
            true
        }
        FillStyle::Gradient {
            kind, stops, angle, ..
        } => {
            if stops.is_empty() {
                return false;
            }
            let mut ordered_stops = stops.clone();
            ordered_stops.sort_by(|left, right| {
                left.position
                    .partial_cmp(&right.position)
                    .unwrap_or(std::cmp::Ordering::Equal)
            });
            let x = rect.x as f64;
            let y = rect.y as f64;
            let w = rect.w as f64;
            let h = rect.h as f64;
            let gradient = if kind == "radial" {
                let radius = w.max(h) / 2.0;
                ctx.create_radial_gradient(
                    x + w / 2.0,
                    y + h / 2.0,
                    0.0,
                    x + w / 2.0,
                    y + h / 2.0,
                    radius,
                )
                .ok()
            } else {
                let radians = angle.unwrap_or(0.0) as f64 * std::f64::consts::PI / 180.0;
                let dx = radians.cos() * w / 2.0;
                let dy = radians.sin() * h / 2.0;
                Some(ctx.create_linear_gradient(
                    x + w / 2.0 - dx,
                    y + h / 2.0 - dy,
                    x + w / 2.0 + dx,
                    y + h / 2.0 + dy,
                ))
            };
            let Some(gradient) = gradient else {
                return false;
            };
            for stop in &ordered_stops {
                let _ = gradient.add_color_stop(stop.position, &stop.color);
            }
            if stroke {
                ctx.set_stroke_style_canvas_gradient(&gradient);
            } else {
                ctx.set_fill_style_canvas_gradient(&gradient);
            }
            true
        }
    }
}

pub fn begin_shape_path(ctx: &CanvasRenderingContext2d, shp: &ShapeElement) {
    let x = shp.rect.x as f64;
    let y = shp.rect.y as f64;
    let w = shp.rect.w as f64;
    let h = shp.rect.h as f64;
    ctx.begin_path();
    match shp.shape_type.as_str() {
        "roundRect" => {
            let radius = (w.min(h) * shp.corner_radius.unwrap_or(1.0 / 6.0).clamp(0.0, 0.5) as f64)
                .min(w.min(h) / 2.0);
            ctx.move_to(x + radius, y);
            ctx.line_to(x + w - radius, y);
            ctx.quadratic_curve_to(x + w, y, x + w, y + radius);
            ctx.line_to(x + w, y + h - radius);
            ctx.quadratic_curve_to(x + w, y + h, x + w - radius, y + h);
            ctx.line_to(x + radius, y + h);
            ctx.quadratic_curve_to(x, y + h, x, y + h - radius);
            ctx.line_to(x, y + radius);
            ctx.quadratic_curve_to(x, y, x + radius, y);
        }
        "ellipse" => {
            let _ = ctx.ellipse(
                x + w / 2.0,
                y + h / 2.0,
                w / 2.0,
                h / 2.0,
                0.0,
                0.0,
                2.0 * std::f64::consts::PI,
            );
        }
        "triangle" => {
            ctx.move_to(x + w / 2.0, y);
            ctx.line_to(x + w, y + h);
            ctx.line_to(x, y + h);
        }
        "mathPlus" => {
            let arm = (w.min(h) / 3.0).max(0.0);
            let left = x + (w - arm) / 2.0;
            let right = left + arm;
            let top = y + (h - arm) / 2.0;
            let bottom = top + arm;
            ctx.move_to(left, y);
            ctx.line_to(right, y);
            ctx.line_to(right, top);
            ctx.line_to(x + w, top);
            ctx.line_to(x + w, bottom);
            ctx.line_to(right, bottom);
            ctx.line_to(right, y + h);
            ctx.line_to(left, y + h);
            ctx.line_to(left, bottom);
            ctx.line_to(x, bottom);
            ctx.line_to(x, top);
            ctx.line_to(left, top);
        }
        "line" => {
            ctx.move_to(x, y);
            ctx.line_to(x + w, y + h);
        }
        _ => ctx.rect(x, y, w, h),
    }
    if shp.shape_type != "line" {
        ctx.close_path();
    }
}

fn apply_shape_transform(ctx: &CanvasRenderingContext2d, shp: &ShapeElement) {
    let center_x = (shp.rect.x + shp.rect.w / 2.0) as f64;
    let center_y = (shp.rect.y + shp.rect.h / 2.0) as f64;
    let _ = ctx.translate(center_x, center_y);
    let _ = ctx.rotate((shp.rotation as f64).to_radians());
    let _ = ctx.scale(
        if shp.flip_h { -1.0 } else { 1.0 },
        if shp.flip_v { -1.0 } else { 1.0 },
    );
    let _ = ctx.translate(-center_x, -center_y);
}

fn apply_line_dash(ctx: &CanvasRenderingContext2d, dash: Option<&str>, width: f32) {
    let Some(dash) = dash else {
        return;
    };
    let unit = width.max(1.0) as f64;
    let pattern: &[f64] = match dash {
        "dot" => &[unit, unit * 3.0],
        // WPS uses a visibly longer preset dash for the 0.5pt border used by
        // these rounded boxes. The short [3, 2] approximation makes the
        // border look almost solid at the slide scale.
        "dash" | "sysDash" => &[unit * 6.0, unit * 4.0],
        "lgDash" => &[unit * 8.0, unit * 4.0],
        "dashDot" | "sysDashDot" => &[unit * 6.0, unit * 4.0, unit, unit * 4.0],
        "lgDashDot" => &[unit * 8.0, unit * 4.0, unit, unit * 4.0],
        "dashDotDot" | "sysDashDotDot" => {
            &[unit * 6.0, unit * 4.0, unit, unit * 4.0, unit, unit * 4.0]
        }
        "lgDashDotDot" => &[unit * 8.0, unit * 4.0, unit, unit * 4.0, unit, unit * 4.0],
        _ => return,
    };
    let values = Array::new();
    for value in pattern {
        values.push(&JsValue::from_f64(*value));
    }
    let _ = ctx.set_line_dash(&values);
}

pub fn paint_shape(ctx: &CanvasRenderingContext2d, shp: &ShapeElement) {
    ctx.save();
    apply_shape_transform(ctx, shp);
    begin_shape_path(ctx, shp);
    if let Some(style) = &shp.computed_style {
        if shp.shape_type != "line" && set_shape_paint(ctx, &style.fill, &shp.rect, false) {
            ctx.fill();
        }
        if let Some(line) = &style.line {
            if set_shape_paint(ctx, &line.fill, &shp.rect, true) {
                ctx.set_line_width(line.width as f64);
                apply_line_dash(ctx, line.dash.as_deref(), line.width);
                if let Some(cap) = &line.cap {
                    ctx.set_line_cap(cap);
                }
                if let Some(join) = &line.join {
                    ctx.set_line_join(join);
                }
                ctx.stroke();
            }
        }
    } else {
        if shp.shape_type != "line" && shp.fill != "transparent" {
            ctx.set_fill_style_str(&shp.fill);
            ctx.fill();
        }
        if let Some(border) = &shp.border {
            ctx.set_stroke_style_str(&border.color);
            ctx.set_line_width(border.width as f64);
            ctx.stroke();
        }
    }
    ctx.restore();
}
