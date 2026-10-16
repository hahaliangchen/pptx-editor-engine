# pptx-editor-engine

A lightweight standalone PPTX rendering and editing engine powered by Rust (WebAssembly) & TypeScript.

## Architecture
- **Parser (TypeScript)**: Parses `.pptx` XML files into a clean JSON Abstract Syntax Tree (AST / Virtual DOM).
- **Layout & Layout Metrics (Rust WASM)**: Computes high-fidelity line-wrapping, fonts, and styling metrics inside the WASM core using `cosmic-text`.
- **Renderer (Canvas)**: Renders the structured AST shapes and text elements smoothly using HTML5 Canvas.

## 检查导出的 PPTX 是否会让 PowerPoint 报“修复”

没有 PowerPoint 时，用微软自己的 Open XML SDK 校验器检查导出的文件（需要 .NET 8，可装在用户目录，不用 sudo）：

```bash
curl -sSL https://dot.net/v1/dotnet-install.sh | bash -s -- --channel 8.0 --install-dir ~/.dotnet
npm run validate:pptx -- 导出的文件.pptx 原始文件.pptx   # 只显示导出后新增的问题
```

测试里另有一套包完整性检查（`tests/helpers/package-check.ts`）：关系目标是否存在、内容类型是否声明、
形状 id 是否重复、幻灯片列表与备注页是否一致。
