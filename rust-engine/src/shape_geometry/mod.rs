//! DrawingML preset geometry dispatch.
//!
//! The parser keeps the original `a:prstGeom@prst` name.  Geometry families
//! live in separate modules so adding a preset does not turn one renderer into
//! an unmaintainable list of unrelated paths.

mod arrows;
mod basic;
mod brackets;
mod callouts;
mod connectors;
mod custom;
mod flowcharts;
mod helpers;
mod polygons;
mod special;
mod stars;
mod symbols;

use crate::ast::ShapeElement;
use wasm_bindgen::JsValue;
use web_sys::CanvasRenderingContext2d;

pub fn draw_fill(ctx: &CanvasRenderingContext2d, shp: &ShapeElement) -> bool {
    draw_mode(ctx, shp, custom::PaintMode::Fill)
}

pub fn draw_stroke(ctx: &CanvasRenderingContext2d, shp: &ShapeElement) -> bool {
    draw_mode(ctx, shp, custom::PaintMode::Stroke)
}

pub fn custom_path_count(shp: &ShapeElement) -> Option<usize> {
    custom::path_count(shp)
}

pub fn draw_custom_path(
    ctx: &CanvasRenderingContext2d,
    shp: &ShapeElement,
    path_index: usize,
) -> bool {
    custom::draw_path(ctx, shp, path_index, custom::PaintMode::Fill)
}

pub fn custom_path_fill_mode(shp: &ShapeElement, path_index: usize) -> Option<&str> {
    custom::path_fill_mode(shp, path_index)
}

fn draw_mode(ctx: &CanvasRenderingContext2d, shp: &ShapeElement, mode: custom::PaintMode) -> bool {
    if shp.custom_geometry.is_some() {
        return custom::draw(ctx, shp, mode);
    }
    let (x, y, w, h) = (
        shp.rect.x as f64,
        shp.rect.y as f64,
        shp.rect.w as f64,
        shp.rect.h as f64,
    );
    if basic::draw(ctx, shp, x, y, w, h)
        || polygons::draw(ctx, shp, x, y, w, h)
        || stars::draw(ctx, shp, x, y, w, h)
        || arrows::draw(ctx, shp, x, y, w, h)
        || connectors::draw(ctx, shp, x, y, w, h)
        || brackets::draw(ctx, shp, x, y, w, h)
        || callouts::draw(ctx, shp, x, y, w, h)
        || flowcharts::draw(ctx, shp, x, y, w, h)
        || symbols::draw(ctx, shp, x, y, w, h)
        || special::draw(ctx, shp, x, y, w, h)
    {
        return true;
    }

    web_sys::console::warn_1(&JsValue::from_str(&format!(
        "[ShapeGeometry] unsupported preset geometry: {}",
        shp.shape_type
    )));
    false
}

pub fn has_fill(shp: &ShapeElement) -> bool {
    if shp.custom_geometry.is_some() {
        return custom::has_fill(shp);
    }
    !is_stroke_only(&shp.shape_type)
}

pub fn is_custom(shp: &ShapeElement) -> bool {
    shp.custom_geometry.is_some()
}

pub fn contains(shp: &ShapeElement, local_x: f32, local_y: f32) -> bool {
    let w = shp.rect.w.max(0.0);
    let h = shp.rect.h.max(0.0);
    if local_x < 0.0 || local_y < 0.0 || local_x > w || local_y > h {
        return false;
    }
    custom::contains(shp, local_x, local_y)
        .or_else(|| basic::contains(shp, local_x, local_y, w, h))
        .or_else(|| polygons::contains(shp, local_x, local_y, w, h))
        .or_else(|| stars::contains(shp, local_x, local_y, w, h))
        .or_else(|| arrows::contains(shp, local_x, local_y, w, h))
        .or_else(|| connectors::contains(shp, local_x, local_y, w, h))
        .or_else(|| callouts::contains(shp, local_x, local_y, w, h))
        .or_else(|| flowcharts::contains(shp, local_x, local_y, w, h))
        .unwrap_or(true)
}

pub fn is_stroke_only(shape_type: &str) -> bool {
    shape_type == "line"
        || shape_type == "lineInv"
        || shape_type == "arc"
        || shape_type == "straightConnector1"
        || shape_type.starts_with("bentConnector")
        || shape_type.starts_with("curvedConnector")
        || matches!(
            shape_type,
            "leftBracket"
                | "rightBracket"
                | "leftBrace"
                | "rightBrace"
                | "bracketPair"
                | "bracePair"
        )
}
