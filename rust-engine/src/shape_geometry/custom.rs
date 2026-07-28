use std::f64::consts::PI;

use crate::ast::{CustomGeometryCommand, ShapeElement};
use web_sys::CanvasRenderingContext2d;

#[derive(Clone, Copy)]
pub enum PaintMode {
    Fill,
    Stroke,
}

fn path_is_included(path: &crate::ast::CustomGeometryPath, mode: PaintMode) -> bool {
    match mode {
        PaintMode::Fill => path.fill.as_deref() != Some("none"),
        PaintMode::Stroke => path.stroke.unwrap_or(true),
    }
}

fn angle_radians(value: f32) -> f64 {
    (value as f64 / 60000.0) * PI / 180.0
}

fn path_scale(shape: &ShapeElement, width: f32, height: f32) -> (f64, f64) {
    (
        shape.rect.w as f64 / width.max(1.0) as f64,
        shape.rect.h as f64 / height.max(1.0) as f64,
    )
}

fn append_arc(
    ctx: &CanvasRenderingContext2d,
    current: &mut Option<(f64, f64)>,
    width_radius: f32,
    height_radius: f32,
    start_angle: f32,
    sweep_angle: f32,
) {
    let start = angle_radians(start_angle);
    let sweep = angle_radians(sweep_angle);
    let rx = (width_radius as f64).abs().max(0.001);
    let ry = (height_radius as f64).abs().max(0.001);
    let (start_x, start_y) = current.unwrap_or((rx * start.cos(), ry * start.sin()));
    let center_x = start_x - rx * start.cos();
    let center_y = start_y - ry * start.sin();
    let _ = ctx.ellipse(center_x, center_y, rx, ry, 0.0, start, start + sweep);
    *current = Some((
        center_x + rx * (start + sweep).cos(),
        center_y + ry * (start + sweep).sin(),
    ));
}

pub fn draw_path(
    ctx: &CanvasRenderingContext2d,
    shape: &ShapeElement,
    path_index: usize,
    mode: PaintMode,
) -> bool {
    let Some(path) = shape
        .custom_geometry
        .as_ref()
        .and_then(|geometry| geometry.paths.get(path_index))
    else {
        return false;
    };
    if path.commands.is_empty() || !path_is_included(path, mode) {
        return false;
    }
    let (sx, sy) = path_scale(shape, path.width, path.height);
    let _ = ctx.save();
    let _ = ctx.translate(shape.rect.x as f64, shape.rect.y as f64);
    let _ = ctx.scale(sx, sy);
    let mut current = None;

    for command in &path.commands {
        match command {
            CustomGeometryCommand::MoveTo { x, y } => {
                ctx.move_to(*x as f64, *y as f64);
                current = Some((*x as f64, *y as f64));
            }
            CustomGeometryCommand::LineTo { x, y } => {
                ctx.line_to(*x as f64, *y as f64);
                current = Some((*x as f64, *y as f64));
            }
            CustomGeometryCommand::QuadBezier {
                control_x,
                control_y,
                x,
                y,
            } => {
                ctx.quadratic_curve_to(*control_x as f64, *control_y as f64, *x as f64, *y as f64);
                current = Some((*x as f64, *y as f64));
            }
            CustomGeometryCommand::CubicBezier {
                control1_x,
                control1_y,
                control2_x,
                control2_y,
                x,
                y,
            } => {
                ctx.bezier_curve_to(
                    *control1_x as f64,
                    *control1_y as f64,
                    *control2_x as f64,
                    *control2_y as f64,
                    *x as f64,
                    *y as f64,
                );
                current = Some((*x as f64, *y as f64));
            }
            CustomGeometryCommand::Arc {
                width_radius,
                height_radius,
                start_angle,
                sweep_angle,
            } => append_arc(
                ctx,
                &mut current,
                *width_radius,
                *height_radius,
                *start_angle,
                *sweep_angle,
            ),
            CustomGeometryCommand::Close => {
                ctx.close_path();
            }
        }
    }
    let _ = ctx.restore();
    true
}

pub fn draw(ctx: &CanvasRenderingContext2d, shape: &ShapeElement, mode: PaintMode) -> bool {
    let Some(geometry) = &shape.custom_geometry else {
        return false;
    };
    let mut drew = false;
    for path_index in 0..geometry.paths.len() {
        drew |= draw_path(ctx, shape, path_index, mode);
    }
    drew
}

pub fn path_count(shape: &ShapeElement) -> Option<usize> {
    shape
        .custom_geometry
        .as_ref()
        .map(|geometry| geometry.paths.len())
}

pub fn path_fill_mode(shape: &ShapeElement, path_index: usize) -> Option<&str> {
    shape
        .custom_geometry
        .as_ref()
        .and_then(|geometry| geometry.paths.get(path_index))
        .and_then(|path| path.fill.as_deref())
}

fn sample_cubic(
    points: &mut Vec<(f64, f64)>,
    start: (f64, f64),
    control1: (f64, f64),
    control2: (f64, f64),
    end: (f64, f64),
) {
    for index in 1..=12 {
        let t = index as f64 / 12.0;
        let u = 1.0 - t;
        points.push((
            u.powi(3) * start.0
                + 3.0 * u.powi(2) * t * control1.0
                + 3.0 * u * t.powi(2) * control2.0
                + t.powi(3) * end.0,
            u.powi(3) * start.1
                + 3.0 * u.powi(2) * t * control1.1
                + 3.0 * u * t.powi(2) * control2.1
                + t.powi(3) * end.1,
        ));
    }
}

fn sample_quad(
    points: &mut Vec<(f64, f64)>,
    start: (f64, f64),
    control: (f64, f64),
    end: (f64, f64),
) {
    for index in 1..=10 {
        let t = index as f64 / 10.0;
        let u = 1.0 - t;
        points.push((
            u * u * start.0 + 2.0 * u * t * control.0 + t * t * end.0,
            u * u * start.1 + 2.0 * u * t * control.1 + t * t * end.1,
        ));
    }
}

fn sample_arc(
    points: &mut Vec<(f64, f64)>,
    start_point: (f64, f64),
    width_radius: f32,
    height_radius: f32,
    start_angle: f32,
    sweep_angle: f32,
) -> (f64, f64) {
    let start = angle_radians(start_angle);
    let sweep = angle_radians(sweep_angle);
    let rx = (width_radius as f64).abs().max(0.001);
    let ry = (height_radius as f64).abs().max(0.001);
    let center = (
        start_point.0 - rx * start.cos(),
        start_point.1 - ry * start.sin(),
    );
    let steps = ((sweep.abs() / (PI / 18.0)).ceil() as usize).clamp(2, 72);
    for index in 1..=steps {
        let angle = start + sweep * index as f64 / steps as f64;
        points.push((center.0 + rx * angle.cos(), center.1 + ry * angle.sin()));
    }
    points.last().copied().unwrap_or(start_point)
}

fn path_points(
    shape: &ShapeElement,
    width: f32,
    height: f32,
    commands: &[CustomGeometryCommand],
) -> Vec<(f64, f64)> {
    let (sx, sy) = path_scale(shape, width, height);
    let mut points = Vec::new();
    let mut current = (0.0, 0.0);
    for command in commands {
        let raw = match command {
            CustomGeometryCommand::MoveTo { x, y } => {
                current = (*x as f64, *y as f64);
                Some(current)
            }
            CustomGeometryCommand::LineTo { x, y } => {
                current = (*x as f64, *y as f64);
                Some(current)
            }
            CustomGeometryCommand::QuadBezier {
                control_x,
                control_y,
                x,
                y,
            } => {
                let end = (*x as f64, *y as f64);
                sample_quad(
                    &mut points,
                    current,
                    (*control_x as f64, *control_y as f64),
                    end,
                );
                current = end;
                None
            }
            CustomGeometryCommand::CubicBezier {
                control1_x,
                control1_y,
                control2_x,
                control2_y,
                x,
                y,
            } => {
                let end = (*x as f64, *y as f64);
                sample_cubic(
                    &mut points,
                    current,
                    (*control1_x as f64, *control1_y as f64),
                    (*control2_x as f64, *control2_y as f64),
                    end,
                );
                current = end;
                None
            }
            CustomGeometryCommand::Arc {
                width_radius,
                height_radius,
                start_angle,
                sweep_angle,
            } => {
                current = sample_arc(
                    &mut points,
                    current,
                    *width_radius,
                    *height_radius,
                    *start_angle,
                    *sweep_angle,
                );
                None
            }
            CustomGeometryCommand::Close => None,
        };
        if let Some(point) = raw {
            points.push(point);
        }
    }
    points.into_iter().map(|(x, y)| (x * sx, y * sy)).collect()
}

fn point_in_polygon(point: (f64, f64), polygon: &[(f64, f64)]) -> bool {
    if polygon.len() < 3 {
        return false;
    }
    let (px, py) = point;
    let mut inside = false;
    let mut previous = *polygon.last().unwrap();
    for &current in polygon {
        let intersects = ((current.1 > py) != (previous.1 > py))
            && (px
                < (previous.0 - current.0) * (py - current.1) / (previous.1 - current.1 + 1e-12)
                    + current.0);
        if intersects {
            inside = !inside;
        }
        previous = current;
    }
    inside
}

pub fn contains(shape: &ShapeElement, local_x: f32, local_y: f32) -> Option<bool> {
    let geometry = shape.custom_geometry.as_ref()?;
    let point = (local_x as f64, local_y as f64);
    Some(geometry.paths.iter().any(|path| {
        path_is_included(path, PaintMode::Fill)
            && point_in_polygon(
                point,
                &path_points(shape, path.width, path.height, &path.commands),
            )
    }))
}

pub fn has_fill(shape: &ShapeElement) -> bool {
    shape
        .custom_geometry
        .as_ref()
        .map(|geometry| {
            geometry
                .paths
                .iter()
                .any(|path| path_is_included(path, PaintMode::Fill))
        })
        .unwrap_or(false)
}

#[cfg(test)]
mod tests {
    use super::{path_is_included, PaintMode};
    use crate::ast::CustomGeometryPath;

    #[test]
    fn path_level_fill_and_stroke_flags_are_independent() {
        let path = CustomGeometryPath {
            width: 100.0,
            height: 100.0,
            commands: Vec::new(),
            fill: Some("none".to_string()),
            stroke: Some(false),
        };
        assert!(!path_is_included(&path, PaintMode::Fill));
        assert!(!path_is_included(&path, PaintMode::Stroke));

        let default_path = CustomGeometryPath {
            width: 100.0,
            height: 100.0,
            commands: Vec::new(),
            fill: None,
            stroke: None,
        };
        assert!(path_is_included(&default_path, PaintMode::Fill));
        assert!(path_is_included(&default_path, PaintMode::Stroke));
    }

    #[test]
    fn arc_sampling_keeps_the_declared_endpoint() {
        let mut points = Vec::new();
        let end = super::sample_arc(&mut points, (10.0, 0.0), 10.0, 10.0, 0.0, 5_400_000.0);
        assert!((end.0 - 0.0).abs() < 0.001);
        assert!((end.1 - 10.0).abs() < 0.001);
    }
}
