import format from "xml-formatter";
import PptViewer from "./index";
import { PptVirtualDocument, Slide } from "./pptx-parser";
import "./style.css";

// Pretty-print XML text using standard 'xml-formatter' library
function prettyPrintXml(xmlString: string, indentString: string = "  "): string {
  if (!xmlString) return "No XML data available for this slide.";
  try {
    return format(xmlString, {
      indentation: indentString,
      collapseContent: true,
      lineSeparator: "\n"
    });
  } catch (err) {
    console.error("xml-formatter failed, returning raw XML string:", err);
    return xmlString;
  }
}

// Setup Mock Presentation Demo with embedded mock XML strings
const mockPresentation: PptVirtualDocument = {
  size: { width: 1920, height: 1080 },
  styleRegistry: { rules: {} },
  slides: [
    {
      id: "slide_1",
      rawXml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr/>
      <p:grpSpPr/>
      <!-- Slide Background Shape -->
      <p:sp>
        <p:spPr>
          <a:xfrm>
            <a:off x="0" y="0"/>
            <a:ext cx="9144000" cy="5142864"/>
          </a:xfrm>
          <a:prstGeom prst="rect"/>
          <a:solidFill>
            <a:srgbClr val="0F172A"/>
          </a:solidFill>
        </p:spPr>
      </p:sp>
      <!-- Title Box Shape and Text -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title Box"/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="952500" y="714375"/>
            <a:ext cx="7239000" cy="3714750"/>
          </a:xfrm>
          <a:prstGeom prst="rect"/>
          <a:solidFill>
            <a:srgbClr val="1E293B"/>
          </a:solidFill>
        </p:spPr>
        <p:txBody>
          <a:bodyPr anchor="ctr"/>
          <a:p>
            <a:pPr algn="ctr"/>
            <a:r>
              <a:rPr sz="6400" b="1">
                <a:solidFill><a:srgbClr val="38BDF8"/></a:solidFill>
              </a:rPr>
              <a:t>PPTX Virtual DOM &amp; WASM Rendering Engine</a:t>
            </a:r>
          </a:p>
          <a:p>
            <a:pPr algn="ctr"/>
            <a:r>
              <a:rPr sz="3200">
                <a:solidFill><a:srgbClr val="94A3B8"/></a:solidFill>
              </a:rPr>
              <a:t>Powered by Rust (WebAssembly) &amp; TypeScript</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`,
      elements: [
        {
          type: "shape",
          id: "bg_1",
          shapeType: "rect",
          rect: { x: 0, y: 0, w: 1920, h: 1080 },
          fill: "#0f172a"
        },
        {
          type: "shape",
          id: "accent_card",
          shapeType: "rect",
          rect: { x: 200, y: 150, w: 1520, h: 780 },
          fill: "#1e293b"
        },
        {
          type: "text",
          id: "title",
          rect: { x: 300, y: 350, w: 1320, h: 200 },
          content: "PPTX Virtual DOM & WASM Rendering Engine",
          style: { fontSize: 64, color: "#38bdf8", bold: true, align: "center" }
        },
        {
          type: "text",
          id: "subtitle",
          rect: { x: 300, y: 550, w: 1320, h: 100 },
          content: "Powered by Rust (WebAssembly) & TypeScript",
          style: { fontSize: 32, color: "#94a3b8", bold: false, align: "center" }
        },
        {
          type: "shape",
          id: "decor_1",
          shapeType: "ellipse",
          rect: { x: 910, y: 700, w: 100, h: 100 },
          fill: "#6366f1"
        }
      ]
    },
    {
      id: "slide_2",
      rawXml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr/>
      <p:grpSpPr/>
      <p:sp>
        <p:spPr>
          <a:xfrm>
            <a:off x="0" y="0"/>
            <a:ext cx="9144000" cy="5142864"/>
          </a:xfrm>
          <a:prstGeom prst="rect"/>
          <a:solidFill><a:srgbClr val="0F172A"/></a:solidFill>
        </p:spPr>
      </p:sp>
      <p:sp>
        <p:txBody>
          <a:p>
            <a:r>
              <a:rPr sz="4400" b="1"><a:solidFill><a:srgbClr val="38BDF8"/></a:solidFill></a:rPr>
              <a:t>Supported Elements &amp; Styles</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <!-- Column 1: Vector shapes -->
      <p:sp>
        <p:spPr>
          <a:xfrm>
            <a:off x="476250" y="1190625"/>
            <a:ext cx="2381250" cy="2857500"/>
          </a:xfrm>
          <a:prstGeom prst="rect"/>
          <a:solidFill><a:srgbClr val="1E293B"/></a:solidFill>
        </p:spPr>
        <p:txBody>
          <a:p>
            <a:r>
              <a:rPr sz="2400"><a:solidFill><a:srgbClr val="E2E8F0"/></a:solidFill></a:rPr>
              <a:t>Vector Shapes&#x0A;&#x0A;- Rectangle&#x0A;- Ellipse / Circle&#x0A;- Triangle&#x0A;&#x0A;Rendered natively in Rust via Canvas.</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`,
      elements: [
        {
          type: "shape",
          id: "bg_2",
          shapeType: "rect",
          rect: { x: 0, y: 0, w: 1920, h: 1080 },
          fill: "#0f172a"
        },
        {
          type: "text",
          id: "slide2_title",
          rect: { x: 100, y: 100, w: 1720, h: 100 },
          content: "Supported Elements & Styles",
          style: { fontSize: 44, color: "#38bdf8", bold: true, align: "left" }
        },
        {
          type: "shape",
          id: "col_1",
          shapeType: "rect",
          rect: { x: 100, y: 250, w: 500, h: 600 },
          fill: "#1e293b"
        },
        {
          type: "text",
          id: "col_1_txt",
          rect: { x: 130, y: 280, w: 440, h: 540 },
          content: "Vector Shapes\n\n- Rectangle\n- Ellipse / Circle\n- Triangle\n\nRendered natively in Rust via HTML5 Canvas context.",
          style: { fontSize: 24, color: "#e2e8f0", bold: false, align: "left" }
        },
        {
          type: "shape",
          id: "col_2",
          shapeType: "rect",
          rect: { x: 710, y: 250, w: 500, h: 600 },
          fill: "#1e293b"
        },
        {
          type: "text",
          id: "col_2_txt",
          rect: { x: 740, y: 280, w: 440, h: 540 },
          content: "Rich Typography\n\n- Customizable Sizing\n- Font Weight (Bold)\n- Hex Colors\n- Alignments (Left, Center, Right)",
          style: { fontSize: 24, color: "#e2e8f0", bold: false, align: "left" }
        },
        {
          type: "shape",
          id: "col_3",
          shapeType: "rect",
          rect: { x: 1320, y: 250, w: 500, h: 600 },
          fill: "#1e293b"
        },
        {
          type: "text",
          id: "col_3_txt",
          rect: { x: 1350, y: 280, w: 440, h: 540 },
          content: "Single Source of Truth\n\nEverything you see is parsed into a JSON Virtual DOM. Extremely flexible for interactive editors, dragging, and resizing.",
          style: { fontSize: 24, color: "#e2e8f0", bold: false, align: "left" }
        }
      ]
    },
    {
      id: "slide_3",
      rawXml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr/>
      <p:grpSpPr/>
      <p:sp>
        <p:spPr>
          <a:xfrm>
            <a:off x="0" y="0"/>
            <a:ext cx="9144000" cy="5142864"/>
          </a:xfrm>
          <a:prstGeom prst="rect"/>
          <a:solidFill><a:srgbClr val="0F172A"/></a:solidFill>
        </p:spPr>
      </p:sp>
      <!-- Geometry shapes -->
      <p:sp>
        <p:spPr>
          <a:xfrm>
            <a:off x="1190625" y="1428750"/>
            <a:ext cx="1428750" cy="1428750"/>
          </a:xfrm>
          <a:prstGeom prst="rect"/>
          <a:solidFill><a:srgbClr val="F43F5E"/></a:solidFill>
        </p:spPr>
      </p:sp>
      <p:sp>
        <p:spPr>
          <a:xfrm>
            <a:off x="3857625" y="1428750"/>
            <a:ext cx="1428750" cy="1428750"/>
          </a:xfrm>
          <a:prstGeom prst="ellipse"/>
          <a:solidFill><a:srgbClr val="10B981"/></a:solidFill>
        </p:spPr>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`,
      elements: [
        {
          type: "shape",
          id: "bg_3",
          shapeType: "rect",
          rect: { x: 0, y: 0, w: 1920, h: 1080 },
          fill: "#0f172a"
        },
        {
          type: "text",
          id: "slide3_title",
          rect: { x: 100, y: 100, w: 1720, h: 100 },
          content: "Theme and Vector Performance Showcase",
          style: { fontSize: 44, color: "#38bdf8", bold: true, align: "left" }
        },
        {
          type: "shape",
          id: "sh_rect",
          shapeType: "rect",
          rect: { x: 250, y: 300, w: 300, h: 300 },
          fill: "#f43f5e"
        },
        {
          type: "shape",
          id: "sh_ellipse",
          shapeType: "ellipse",
          rect: { x: 810, y: 300, w: 300, h: 300 },
          fill: "#10b981"
        },
        {
          type: "shape",
          id: "sh_triangle",
          shapeType: "triangle",
          rect: { x: 1370, y: 300, w: 300, h: 300 },
          fill: "#f59e0b"
        },
        {
          type: "text",
          id: "labels",
          rect: { x: 100, y: 700, w: 1720, h: 200 },
          content: "Natively compiled to WebAssembly for smooth 60fps rendering pipelines.",
          style: { fontSize: 28, color: "#94a3b8", bold: false, align: "center" }
        }
      ]
    }
  ]
};

document.addEventListener("DOMContentLoaded", () => {
  const container = document.getElementById("canvas-wrapper");
  const welcomeView = document.getElementById("welcome-view");
  const pageIndicator = document.getElementById("page-indicator");
  const btnPrev = document.getElementById("btn-prev") as HTMLButtonElement;
  const btnNext = document.getElementById("btn-next") as HTMLButtonElement;
  const slideList = document.getElementById("slide-list");
  
  // Inspectors
  const astInspector = document.getElementById("ast-inspector");
  const xmlInspector = document.getElementById("xml-inspector");
  
  // Tab buttons
  const tabBtnJson = document.getElementById("tab-btn-json") as HTMLButtonElement;
  const tabBtnXml = document.getElementById("tab-btn-xml") as HTMLButtonElement;

  if (!container || !welcomeView || !pageIndicator || !btnPrev || !btnNext || !slideList || !astInspector || !xmlInspector || !tabBtnJson || !tabBtnXml) {
    console.error("Missing required DOM elements");
    return;
  }

  // Hook up Inspector Tab Toggle
  tabBtnJson.onclick = () => {
    tabBtnJson.classList.add("active");
    tabBtnJson.style.color = "#3b82f6";
    tabBtnJson.style.borderBottomColor = "#3b82f6";
    
    tabBtnXml.classList.remove("active");
    tabBtnXml.style.color = "#6b7280";
    tabBtnXml.style.borderBottomColor = "transparent";
    
    astInspector.style.display = "block";
    xmlInspector.style.display = "none";
  };

  tabBtnXml.onclick = () => {
    tabBtnXml.classList.add("active");
    tabBtnXml.style.color = "#3b82f6";
    tabBtnXml.style.borderBottomColor = "#3b82f6";
    
    tabBtnJson.classList.remove("active");
    tabBtnJson.style.color = "#6b7280";
    tabBtnJson.style.borderBottomColor = "transparent";
    
    xmlInspector.style.display = "block";
    astInspector.style.display = "none";
  };

  // Initialize Viewer
  const viewer = new PptViewer({
    container,
    onSlideChange: (index, ast) => {
      // Update page indicator
      pageIndicator.textContent = `Slide ${index + 1} / ${viewer.getSlidesCount()}`;
      
      // Update button states
      btnPrev.disabled = index === 0;
      btnNext.disabled = index === viewer.getSlidesCount() - 1;

      // Highlight sidebar active item
      const items = slideList.querySelectorAll(".slide-item");
      items.forEach((item, idx) => {
        if (idx === index) {
          item.classList.add("active");
        } else {
          item.classList.remove("active");
        }
      });

      // Display the current slide Virtual DOM (excluding rawXml to keep it clean).
      const cleanAst = { ...ast };
      delete (cleanAst as any).rawXml;
      astInspector.textContent = JSON.stringify(cleanAst, null, 2);
      
      // Display current slide formatted XML data
      tabBtnXml.textContent = `Raw XML (${ast.id}.xml)`;
      xmlInspector.textContent = prettyPrintXml(ast.rawXml || "");
    },
    onLoadComplete: (ast) => {
      // Hide welcome view
      welcomeView.style.display = "none";
      
      // Populate Slide list sidebar
      slideList.innerHTML = "";
      ast.slides.forEach((slide, idx) => {
        const li = document.createElement("li");
        li.className = `slide-item ${idx === 0 ? 'active' : ''}`;
        li.onclick = () => viewer.selectSlide(idx);

        const thumb = document.createElement("div");
        thumb.className = "slide-thumb";
        thumb.textContent = (idx + 1).toString();

        const label = document.createElement("div");
        label.textContent = `Slide ${idx + 1}`;
        label.style.fontSize = "14px";
        label.style.fontWeight = "500";

        li.appendChild(thumb);
        li.appendChild(label);
        slideList.appendChild(li);
      });
    }
  });

  // Load Demo triggers
  const loadDemo = () => {
    viewer.loadVirtualDocument(mockPresentation);
  };

  const btnLoadDemo = document.getElementById("btn-load-demo");
  if (btnLoadDemo) btnLoadDemo.onclick = loadDemo;

  const btnWelcomeDemo = document.getElementById("btn-welcome-demo");
  if (btnWelcomeDemo) btnWelcomeDemo.onclick = loadDemo;

  // Controls
  btnPrev.onclick = () => viewer.prevSlide();
  btnNext.onclick = () => viewer.nextSlide();

  // File Uploader
  const fileUploader = document.getElementById("file-uploader") as HTMLInputElement;
  if (fileUploader) {
    fileUploader.onchange = async (e) => {
      const file = (e.target as HTMLInputElement).files?.[0];
      if (!file) return;

      const reader = new FileReader();
      reader.onload = async (evt) => {
        try {
          if (evt.target?.result instanceof ArrayBuffer) {
            await viewer.loadPptx(evt.target.result);
          }
        } catch (err: any) {
          alert(`Error loading PPTX: ${err.message}`);
        }
      };
      reader.readAsArrayBuffer(file);
    };
  }

  // Drag & Drop
  const dropzone = document.getElementById("dropzone");
  if (dropzone) {
    window.addEventListener("dragenter", (e) => {
      e.preventDefault();
      dropzone.classList.add("dragover");
    });

    dropzone.addEventListener("dragleave", (e) => {
      e.preventDefault();
      dropzone.classList.remove("dragover");
    });

    dropzone.addEventListener("dragover", (e) => {
      e.preventDefault();
    });

    dropzone.addEventListener("drop", async (e) => {
      e.preventDefault();
      dropzone.classList.remove("dragover");

      const file = e.dataTransfer?.files[0];
      if (file && file.name.endsWith(".pptx")) {
        const reader = new FileReader();
        reader.onload = async (evt) => {
          try {
            if (evt.target?.result instanceof ArrayBuffer) {
              await viewer.loadPptx(evt.target.result);
            }
          } catch (err: any) {
            alert(`Error loading PPTX: ${err.message}`);
          }
        };
        reader.readAsArrayBuffer(file);
      }
    });
  }
});
