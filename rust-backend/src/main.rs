use axum::{
    extract::{Query, State},
    http::{HeaderMap, HeaderValue, StatusCode},
    response::IntoResponse,
    routing::get,
    Json, Router,
};
use fontdb::{Database, Source};
use serde::Deserialize;
use std::collections::HashSet;
use std::sync::Arc;
use tokio::fs;
use tower_http::cors::{Any, CorsLayer};

struct AppState {
    db: Database,
}

#[tokio::main]
async fn main() {
    let mut db = Database::new();
    println!("Scanning system fonts...");
    db.load_system_fonts();
    println!("Scan complete. Found {} font faces.", db.faces().count());

    let shared_state = Arc::new(AppState { db });

    // Configure CORS so our frontend on localhost:3000 can fetch fonts
    let cors = CorsLayer::new()
        .allow_origin(Any)
        .allow_methods(Any)
        .allow_headers(Any);

    let app = Router::new()
        .route("/api/fonts", get(get_fonts))
        .route("/api/font", get(get_font_file))
        .layer(cors)
        .with_state(shared_state);

    let listener = tokio::net::TcpListener::bind("127.0.0.1:8080")
        .await
        .unwrap();
    println!("Rust Font Backend running on http://127.0.0.1:8080");
    axum::serve(listener, app).await.unwrap();
}

async fn get_fonts(State(state): State<Arc<AppState>>) -> Json<Vec<String>> {
    let mut families = HashSet::new();
    for face in state.db.faces() {
        for (name, _lang) in &face.families {
            families.insert(name.clone());
        }
    }
    let mut families_vec: Vec<String> = families.into_iter().collect();
    families_vec.sort();
    Json(families_vec)
}

#[derive(Deserialize)]
struct FontQuery {
    family: String,
    weight: Option<u16>,
    italic: Option<bool>,
}

async fn get_font_file(
    State(state): State<Arc<AppState>>,
    Query(query): Query<FontQuery>,
) -> impl IntoResponse {
    let q = fontdb::Query {
        families: &[fontdb::Family::Name(&query.family)],
        weight: fontdb::Weight(query.weight.unwrap_or(400)),
        style: if query.italic.unwrap_or(false) {
            fontdb::Style::Italic
        } else {
            fontdb::Style::Normal
        },
        ..Default::default()
    };

    let mut headers = HeaderMap::new();
    headers.insert(
        "Cache-Control",
        HeaderValue::from_static("public, max-age=3600"),
    );

    if let Some(id) = state.db.query(&q) {
        if let Some(face) = state.db.face(id) {
            let path_opt = match &face.source {
                Source::File(path) => Some(path.clone()),
                Source::SharedFile(path, _) => Some(path.clone()),
                _ => None,
            };

            if let Some(path) = path_opt {
                let content_type = match path.extension().and_then(|extension| extension.to_str()) {
                    Some("otf") => "font/otf",
                    Some("ttc") | Some("otc") => "font/collection",
                    Some("woff") => "font/woff",
                    Some("woff2") => "font/woff2",
                    _ => "font/ttf",
                };
                headers.insert("Content-Type", HeaderValue::from_static(content_type));
                match fs::read(&path).await {
                    Ok(bytes) => return (StatusCode::OK, headers, bytes).into_response(),
                    Err(e) => {
                        eprintln!("Failed to read font file at {:?}: {}", path, e);
                    }
                }
            } else if let Source::Binary(bytes) = &face.source {
                // If it's a binary source in the DB, return it directly
                headers.insert("Content-Type", HeaderValue::from_static("font/ttf"));
                let data = bytes.as_ref().as_ref().to_vec();
                return (StatusCode::OK, headers, data).into_response();
            }
        }
    }

    (StatusCode::NOT_FOUND, "Font not found").into_response()
}
