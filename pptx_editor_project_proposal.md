# 项目创建方案：极简 Web PPTX 编辑器与渲染引擎 (NPM 包)

本项目是一个独立的、高性能 Web 端 PPTX 编辑器与渲染引擎。项目被设计为一个独立的 NPM 包，采用 **Rust (WASM) 负责底层计算与渲染，TypeScript 负责事件交互与 DOM 控件** 的黄金双层架构。项目使用 **Rspack + SWC** 作为打包与转译工具链，实现极速的本地构建体验。

---

## 🛠️ 技术栈规范

1. **核心计算与渲染（Rust - WebAssembly）**：
   - 依赖项：`wasm-bindgen`（JS/Rust 绑定）、`serde`/`serde-json`（数据解析）、`cosmic-text`（文本度量与排版折行）。
   - 渲染接口：通过 `web-sys` 直接向 HTML5 Canvas 2D 上下文绘图。
2. **交互与外壳（TypeScript）**：
   - 监听鼠标事件，处理位置移动、8 点选框拖动缩放，并以透明 `<textarea>` 覆盖层解决中文输入法 (IME) 输入问题。
3. **构建与打包工具（Rust 驱动的前端生态）**：
   - **Rspack**：作为 Webpack 的 Rust 高性能替代品，处理打包并原生支持 WASM 模块。
   - **SWC**：作为 Babel 和 TSC 的 Rust 替代品，执行极速的 TS 编译转译。
   - **wasm-pack**：用于将 Rust 编译为符合 NPM 导入规范的 Web/Bundler 格式包。

---

## 📂 项目目录结构蓝图

```text
pptx-editor-engine/
├── package.json               # 定义 NPM 包元数据、依赖和构建脚本
├── tsconfig.json              # TypeScript 编译配置
├── rspack.config.js           # Rspack 打包及库输出配置
├── rust-engine/               # Rust 排版渲染内核 Crate
│   ├── Cargo.toml             # Rust 依赖声明
│   └── src/
│       ├── lib.rs             # WASM 导出函数及初始化入口
│       ├── ast.rs             # PPT AST 语义数据结构
│       └── render.rs          # Canvas 2D 渲染实现
├── src/                       # TypeScript 交互外壳源码
│   ├── index.ts               # NPM 库的统一入口 (PptEditor 类)
│   ├── canvas.ts              # 浏览器事件派发与 Canvas 上下文管理
│   ├── ui.ts                  # 交互把手绘制与行内编辑 Overlay
│   └── style.css              # 编辑器极简 CSS 样式
└── index.html                 # 本地开发调试预览页面
```

---

## ⚙️ 核心配置文件定义

### 1. `package.json`
```json
{
  "name": "@your-org/pptx-editor",
  "version": "1.0.0",
  "description": "A lightweight standalone PPTX rendering and editing engine powered by Rust WASM and TypeScript.",
  "main": "dist/index.js",
  "types": "dist/index.d.ts",
  "files": [
    "dist"
  ],
  "scripts": {
    "build:wasm": "wasm-pack build rust-engine --target web --out-dir pkg",
    "dev": "npm run build:wasm && rspack dev",
    "build": "npm run build:wasm && rspack build"
  },
  "dependencies": {},
  "devDependencies": {
    "@rspack/core": "^0.5.0",
    "@rspack/cli": "^0.5.0",
    "builtin:swc-loader": "^0.5.0",
    "typescript": "^5.0.0",
    "style-loader": "^3.3.3",
    "css-loader": "^6.8.1"
  }
}
```

### 2. `tsconfig.json`
```json
{
  "compilerOptions": {
    "target": "ES2020",
    "module": "ESNext",
    "moduleResolution": "node",
    "strict": true,
    "esModuleInterop": true,
    "skipLibCheck": true,
    "forceConsistentCasingInFileNames": true,
    "declaration": true,
    "outDir": "./dist"
  },
  "include": ["src/**/*"]
}
```

### 3. `rspack.config.js`
```javascript
const path = require("path");

module.exports = {
  entry: "./src/index.ts",
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "index.js",
    library: {
      name: "PptEditor",
      type: "umd",
      export: "default",
    },
    globalObject: "globalThis",
    clean: true,
  },
  resolve: {
    extensions: [".ts", ".js", ".wasm"],
  },
  module: {
    rules: [
      {
        test: /\.ts$/,
        exclude: [/node_modules/],
        use: {
          loader: "builtin:swc-loader",
          options: {
            jsc: {
              parser: {
                syntax: "typescript",
              },
            },
          },
        },
        type: "javascript/auto",
      },
      {
        test: /\.css$/,
        use: ["style-loader", "css-loader"],
      },
    ],
  },
  experiments: {
    asyncWebAssembly: true, // 启用 Rspack 原生异步 WASM 支持
  },
  devServer: {
    port: 3000,
    hot: true,
  },
};
```

---

## 🦀 Rust-Engine 内核骨架

### 1. `rust-engine/Cargo.toml`
```toml
[package]
name = "ppt-engine"
version = "0.1.0"
edition = "2021"

[lib]
crate-type = ["cdylib", "rlib"]

[dependencies]
wasm-bindgen = "0.2.88"
serde = { version = "1.0", features = ["derive"] }
serde_json = "1.0"
cosmic-text = "0.9" # 用于高性能文本排版
percent-encoding = "2.3"

[dependencies.web-sys]
version = "0.3.65"
features = [
  "CanvasRenderingContext2d",
  "HtmlCanvasElement",
  "console"
]
```

### 2. `rust-engine/src/ast.rs`
定义标准的 PPT 语义 AST。
```rust
use serde::{Deserialize, Serialize};

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct Presentation {
    pub size: Size,
    pub slides: Vec<Slide>,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct Size {
    pub width: f32,
    pub height: f32,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct Slide {
    pub id: String,
    pub elements: Vec<Element>,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(tag = "type")]
pub enum Element {
    #[serde(rename = "text")]
    Text(TextElement),
    #[serde(rename = "shape")]
    Shape(ShapeElement),
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct Rect {
    pub x: f32,
    pub y: f32,
    pub w: f32,
    pub h: f32,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct TextElement {
    pub id: String,
    pub rect: Rect,
    pub content: String,
    pub style: TextStyle,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct TextStyle {
    #[serde(rename = "fontSize")]
    pub font_size: f32,
    pub color: String,
    pub bold: bool,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct ShapeElement {
    pub id: String,
    pub rect: Rect,
    #[serde(rename = "shapeType")]
    pub shape_type: String, // "rect" | "ellipse"
    pub fill: String,
}
```

### 3. `rust-engine/src/lib.rs`
WASM 导出入口函数。
```rust
mod ast;

use wasm_bindgen::prelude::*;
use web_sys::CanvasRenderingContext2d;
use ast::{Presentation, Slide};

#[wasm_bindgen]
pub struct RustPptRenderer {
    ctx: CanvasRenderingContext2d,
}

#[wasm_bindgen]
impl RustPptRenderer {
    #[wasm_bindgen(constructor)]
    pub fn new(ctx: CanvasRenderingContext2d) -> Self {
        Self { ctx }
    }

    // 全量解析并绘制单页 Slide
    #[wasm_bindgen]
    pub fn render_slide(&self, slide_json: &str) -> Result<(), JsValue> {
        let slide: Slide = serde_json::from_str(slide_json)
            .map_err(|e| JsValue::from_str(&format!("JSON Parse Error: {}", e)))?;

        // 1. 清空画布
        self.ctx.clear_rect(0.0, 0.0, 1920.0, 1080.0);

        // 2. 从底到顶顺序绘制每个元素 (Z-Index)
        for element in &slide.elements {
            match element {
                ast::Element::Text(txt) => {
                    self.ctx.set_fill_style(&JsValue::from_str(&txt.style.color));
                    self.ctx.set_font(&format!("{}px sans-serif", txt.style.font_size));
                    self.ctx.fill_text(&txt.content, txt.rect.x as f64, (txt.rect.y + txt.style.font_size) as f64)?;
                }
                ast::Element::Shape(shp) => {
                    self.ctx.set_fill_style(&JsValue::from_str(&shp.fill));
                    if shp.shape_type == "rect" {
                        self.ctx.fill_rect(
                            shp.rect.x as f64,
                            shp.rect.y as f64,
                            shp.rect.w as f64,
                            shp.rect.h as f64,
                        );
                    } else if shp.shape_type == "ellipse" {
                        self.ctx.begin_path();
                        self.ctx.ellipse(
                            (shp.rect.x + shp.rect.w / 2.0) as f64,
                            (shp.rect.y + shp.rect.h / 2.0) as f64,
                            (shp.rect.w / 2.0) as f64,
                            (shp.rect.h / 2.0) as f64,
                            0.0,
                            0.0,
                            2.0 * std::f64::consts::PI,
                        )?;
                        self.ctx.fill();
                    }
                }
            }
        }
        Ok(())
    }

    // 辅助工具方法：计算指定边界下最多可容纳的字符数
    #[wasm_bindgen]
    pub fn calculate_max_chars(&self, w: f32, h: f32, font_size: f32, _text_sample: &str) -> usize {
        // 此处可接入 cosmic-text 进行高精度文字包围盒容量计算
        // 临时兜底估算算法：
        let area = w * h;
        let char_area = font_size * (font_size * 0.6); // 估算单个字符的物理面积
        if char_area > 0.0 {
            (area / char_area) as usize
        } else {
            0
        }
    }
}
```

---

## 🖥️ TypeScript 交互壳骨架

### 1. `src/index.ts`
NPM 包的实例化入口。
```typescript
import * as wasm from "../rust-engine/pkg/ppt_engine";

export interface EditorOptions {
  container: HTMLElement;
  width: number;
  height: number;
}

export default class PptEditor {
  private container: HTMLElement;
  private canvas: HTMLCanvasElement;
  private ctx: CanvasRenderingContext2D;
  private renderer!: wasm.RustPptRenderer;
  private ast: any;

  constructor(options: EditorOptions) {
    this.container = options.container;
    
    // 创建 Canvas DOM
    this.canvas = document.createElement("canvas");
    this.canvas.width = options.width;
    this.canvas.height = options.height;
    this.canvas.style.border = "1px solid #ccc";
    this.container.appendChild(this.canvas);
    
    const context = this.canvas.getContext("2d");
    if (!context) throw new Error("Failed to get 2D context");
    this.ctx = context;

    this.initEngine();
  }

  private async initEngine() {
    // 异步加载 Rust WASM 内核
    const wasmModule = await import("../rust-engine/pkg/ppt_engine");
    this.renderer = new wasmModule.RustPptRenderer(this.ctx);
    console.log("PPT WASM Engine initialized.");
  }

  // 挂载新的 slide 数据结构并全量绘制
  public loadSlide(slideJson: any) {
    this.ast = slideJson;
    this.draw();
  }

  private draw() {
    if (!this.renderer || !this.ast) return;
    this.renderer.render_slide(JSON.stringify(this.ast));
  }

  // 命令式数据改变，更新位置后，全量刷新重绘
  public updateElementPosition(id: string, x: number, y: number) {
    const el = this.ast.elements.find((item: any) => item.id === id);
    if (el) {
      el.rect.x = x;
      el.rect.y = y;
      this.draw(); // 全量重绘唯一数据源
    }
  }
}
```

### 2. `index.html` (本地测试预览页)
```html
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>PPT WASM Editor Debug Page</title>
  <style>
    body { display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100vh; margin: 0; background: #fafafa; }
    #editor-container { box-shadow: 0 4px 12px rgba(0,0,0,0.15); background: white; }
  </style>
</head>
<body>
  <h2>PPTX WASM Editor Sandbox</h2>
  <div id="editor-container"></div>

  <script type="module">
    import PptEditor from "./src/index.ts";

    const container = document.getElementById("editor-container");
    const editor = new PptEditor({ container, width: 960, height: 540 });

    const mockSlide = {
      id: "slide_1",
      elements: [
        {
          type: "shape",
          id: "shape_1",
          shapeType: "rect",
          rect: { x: 100, y: 100, w: 200, h: 150 },
          fill: "#1A73E8"
        },
        {
          type: "text",
          id: "text_1",
          rect: { x: 150, y: 300, w: 300, h: 50 },
          content: "Hello from Rust WASM!",
          style: { fontSize: 32, color: "#333", bold: true }
        }
      ]
    };

    setTimeout(() => {
      editor.loadSlide(mockSlide);
    }, 1000);
  </script>
</body>
</html>
```

---

## 🛠️ 初始化与启动说明

若要从零开始创建该项目，请在选定工作目录下顺序运行以下终端指令：

```bash
# 1. 创建项目目录结构
mkdir pptx-editor-engine && cd pptx-editor-engine
mkdir -p src rust-engine/src

# 2. 安装 Node 依赖包
npm install

# 3. 编译 Rust WebAssembly 模块 (确保已安装 rust/wasm-pack)
npm run build:wasm

# 4. 启动本地 Rspack 开发沙箱调试服务
npm run dev
```
