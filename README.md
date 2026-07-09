# pptx-editor-engine

A lightweight standalone PPTX rendering and editing engine powered by Rust (WebAssembly) & TypeScript.

## Architecture
- **Parser (TypeScript)**: Parses `.pptx` XML files into a clean JSON Abstract Syntax Tree (AST / Virtual DOM).
- **Layout & Layout Metrics (Rust WASM)**: Computes high-fidelity line-wrapping, fonts, and styling metrics inside the WASM core using `cosmic-text`.
- **Renderer (Canvas)**: Renders the structured AST shapes and text elements smoothly using HTML5 Canvas.
