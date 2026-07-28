use crate::ast::{FillStyle, Rect, ShapeElement};
use crate::shape_geometry;
use js_sys::Array;
use wasm_bindgen::{JsCast, JsValue};
use web_sys::{CanvasRenderingContext2d, HtmlImageElement};

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
        FillStyle::Pattern { foreground, .. } => {
            // The pattern definition is preserved in the AST. Canvas has no
            // portable preset-pattern primitive, so use the foreground color
            // as a deterministic fallback until a tiled pattern surface is
            // introduced for this renderer.
            if stroke {
                ctx.set_stroke_style_str(foreground);
            } else {
                ctx.set_fill_style_str(foreground);
            }
            true
        }
        FillStyle::Picture { .. } => false,
    }
}

fn begin_shape_path(ctx: &CanvasRenderingContext2d, shp: &ShapeElement, fill: bool) {
    ctx.begin_path();
    let drawn = if fill {
        shape_geometry::draw_fill(ctx, shp)
    } else {
        shape_geometry::draw_stroke(ctx, shp)
    };
    if !drawn {
        ctx.rect(
            shp.rect.x as f64,
            shp.rect.y as f64,
            shp.rect.w as f64,
            shp.rect.h as f64,
        );
    } else if !shape_geometry::is_stroke_only(&shp.shape_type) && !shape_geometry::is_custom(shp) {
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

fn line_endpoints(shp: &ShapeElement) -> Option<((f64, f64), (f64, f64))> {
    let x = shp.rect.x as f64;
    let y = shp.rect.y as f64;
    let right = x + shp.rect.w as f64;
    let bottom = y + shp.rect.h as f64;
    match shp.shape_type.as_str() {
        "line" | "straightConnector1" => Some(((x, y), (right, bottom))),
        "lineInv" => Some(((right, y), (x, bottom))),
        _ => None,
    }
}

fn marker_scale(value: Option<&String>) -> f64 {
    match value.map(String::as_str) {
        Some("sm") => 0.75,
        Some("lg") => 1.5,
        _ => 1.0,
    }
}

fn draw_line_end(
    ctx: &CanvasRenderingContext2d,
    endpoint: (f64, f64),
    outward: (f64, f64),
    end: &crate::ast::LineEndStyle,
    line_width: f32,
) {
    if end.end_type == "none" {
        return;
    }
    let length = line_width.max(1.0) as f64 * 4.0 * marker_scale(end.length.as_ref());
    let half_width = line_width.max(1.0) as f64 * 2.5 * marker_scale(end.width.as_ref());
    let (ux, uy) = outward;
    let (nx, ny) = (-uy, ux);
    let base = (endpoint.0 - ux * length, endpoint.1 - uy * length);
    let left = (base.0 + nx * half_width, base.1 + ny * half_width);
    let right = (base.0 - nx * half_width, base.1 - ny * half_width);

    ctx.begin_path();
    match end.end_type.as_str() {
        "open" => {
            ctx.move_to(left.0, left.1);
            ctx.line_to(endpoint.0, endpoint.1);
            ctx.line_to(right.0, right.1);
            ctx.stroke();
        }
        "diamond" => {
            let rear = (base.0 - ux * length * 0.45, base.1 - uy * length * 0.45);
            ctx.move_to(endpoint.0, endpoint.1);
            ctx.line_to(left.0, left.1);
            ctx.line_to(rear.0, rear.1);
            ctx.line_to(right.0, right.1);
            ctx.close_path();
            ctx.fill();
        }
        "oval" => {
            let _ = ctx.arc(
                endpoint.0,
                endpoint.1,
                half_width,
                0.0,
                std::f64::consts::TAU,
            );
            ctx.fill();
        }
        _ => {
            // triangle, stealth and unknown Office line-end presets use the
            // same conservative filled arrowhead until stealth notch geometry
            // is added.
            ctx.move_to(endpoint.0, endpoint.1);
            ctx.line_to(left.0, left.1);
            ctx.line_to(right.0, right.1);
            ctx.close_path();
            ctx.fill();
        }
    }
}

fn draw_line_ends(
    ctx: &CanvasRenderingContext2d,
    shp: &ShapeElement,
    line: &crate::ast::LineStyle,
) {
    let Some(((start_x, start_y), (end_x, end_y))) = line_endpoints(shp) else {
        return;
    };
    let dx = end_x - start_x;
    let dy = end_y - start_y;
    let length = (dx * dx + dy * dy).sqrt();
    if length <= 0.001 {
        return;
    }
    let ux = dx / length;
    let uy = dy / length;
    if let Some(head) = &line.head_end {
        draw_line_end(ctx, (start_x, start_y), (-ux, -uy), head, line.width);
    }
    if let Some(tail) = &line.tail_end {
        draw_line_end(ctx, (end_x, end_y), (ux, uy), tail, line.width);
    }
}

fn draw_picture_fill(
    ctx: &CanvasRenderingContext2d,
    shp: &ShapeElement,
    fill: &FillStyle,
    images_obj: &JsValue,
) -> bool {
    let FillStyle::Picture {
        url: Some(url),
        src_rect,
        ..
    } = fill
    else {
        return false;
    };
    let Ok(image_value) = js_sys::Reflect::get(images_obj, &JsValue::from_str(url)) else {
        return false;
    };
    let Ok(image) = image_value.dyn_into::<HtmlImageElement>() else {
        return false;
    };
    ctx.save();
    ctx.clip();
    if let Some(crop) = src_rect {
        let source_width = image.natural_width() as f64;
        let source_height = image.natural_height() as f64;
        let sx = source_width * crop.left.clamp(0.0, 1.0) as f64;
        let sy = source_height * crop.top.clamp(0.0, 1.0) as f64;
        let sw = source_width * (1.0 - crop.left - crop.right).clamp(0.0, 1.0) as f64;
        let sh = source_height * (1.0 - crop.top - crop.bottom).clamp(0.0, 1.0) as f64;
        if sw > 0.0 && sh > 0.0 {
            let _ = ctx
                .draw_image_with_html_image_element_and_sw_and_sh_and_dx_and_dy_and_dw_and_dh(
                    &image,
                    sx,
                    sy,
                    sw,
                    sh,
                    shp.rect.x as f64,
                    shp.rect.y as f64,
                    shp.rect.w as f64,
                    shp.rect.h as f64,
                );
        }
    } else {
        let _ = ctx.draw_image_with_html_image_element_and_dw_and_dh(
            &image,
            shp.rect.x as f64,
            shp.rect.y as f64,
            shp.rect.w as f64,
            shp.rect.h as f64,
        );
    }
    ctx.restore();
    true
}

fn apply_custom_fill_mode(ctx: &CanvasRenderingContext2d, mode: Option<&str>) {
    match mode {
        Some("lighten") => {
            let _ = ctx.set_global_composite_operation("screen");
        }
        Some("lightenLess") => {
            let _ = ctx.set_global_composite_operation("screen");
            ctx.set_global_alpha(0.5);
        }
        Some("darken") => {
            let _ = ctx.set_global_composite_operation("multiply");
        }
        Some("darkenLess") => {
            let _ = ctx.set_global_composite_operation("multiply");
            ctx.set_global_alpha(0.5);
        }
        _ => {}
    }
}

fn paint_custom_fills(
    ctx: &CanvasRenderingContext2d,
    shp: &ShapeElement,
    fill: &FillStyle,
    images_obj: &JsValue,
) -> bool {
    let Some(path_count) = shape_geometry::custom_path_count(shp) else {
        return false;
    };
    let mut painted_path = false;
    for path_index in 0..path_count {
        let path_fill_mode = shape_geometry::custom_path_fill_mode(shp, path_index);
        if path_fill_mode == Some("none") {
            continue;
        }
        ctx.begin_path();
        if !shape_geometry::draw_custom_path(ctx, shp, path_index) {
            continue;
        }
        ctx.save();
        apply_custom_fill_mode(ctx, path_fill_mode);
        let picture_filled = matches!(fill, FillStyle::Picture { .. })
            && draw_picture_fill(ctx, shp, fill, images_obj);
        if !picture_filled && set_shape_paint(ctx, fill, &shp.rect, false) {
            ctx.fill();
        }
        ctx.restore();
        painted_path = true;
    }
    painted_path
}

pub fn paint_shape(ctx: &CanvasRenderingContext2d, shp: &ShapeElement, images_obj: &JsValue) {
    ctx.save();
    apply_shape_transform(ctx, shp);
    if let Some(style) = &shp.computed_style {
        if shape_geometry::is_custom(shp) {
            paint_custom_fills(ctx, shp, &style.fill, images_obj);
        } else if shape_geometry::has_fill(shp) {
            begin_shape_path(ctx, shp, true);
            let picture_filled = matches!(style.fill, FillStyle::Picture { .. })
                && draw_picture_fill(ctx, shp, &style.fill, images_obj);
            if !picture_filled && set_shape_paint(ctx, &style.fill, &shp.rect, false) {
                ctx.fill();
            }
        }
        if let Some(line) = &style.line {
            begin_shape_path(ctx, shp, false);
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
                if line.head_end.is_some() || line.tail_end.is_some() {
                    if let FillStyle::Solid { color } = &line.fill {
                        ctx.set_fill_style_str(color);
                        ctx.set_stroke_style_str(color);
                    }
                    draw_line_ends(ctx, shp, line);
                }
            }
        }
    } else {
        if shape_geometry::is_custom(shp) && shp.fill != "transparent" {
            let fallback_fill = FillStyle::Solid {
                color: shp.fill.clone(),
            };
            paint_custom_fills(ctx, shp, &fallback_fill, images_obj);
        } else if shape_geometry::has_fill(shp) && shp.fill != "transparent" {
            begin_shape_path(ctx, shp, true);
            ctx.set_fill_style_str(&shp.fill);
            ctx.fill();
        }
        if let Some(border) = &shp.border {
            begin_shape_path(ctx, shp, false);
            ctx.set_stroke_style_str(&border.color);
            ctx.set_line_width(border.width as f64);
            ctx.stroke();
        }
    }
    ctx.restore();
}
