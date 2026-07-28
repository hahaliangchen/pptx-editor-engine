import {
  CustomGeometry,
  CustomGeometryCommand,
  CustomGeometryPath
} from "./pptx-virtual-dom";
import { getDirectChild, getDirectChildren } from "./pptx-xml";

type GuideTable = Record<string, number>;

function numberOrGuide(value: string | null, guides: GuideTable): number | undefined {
  if (!value) return undefined;
  const numeric = Number(value);
  if (Number.isFinite(numeric)) return numeric;
  const guideValue = guides[value.trim()];
  return guideValue !== undefined && Number.isFinite(guideValue) ? guideValue : undefined;
}

function angleRadians(value: number): number {
  return (value / 60000) * (Math.PI / 180);
}

function evaluateFormula(formula: string | null, guides: GuideTable): number | undefined {
  if (!formula) return undefined;
  const tokens = formula.trim().split(/\s+/).filter(Boolean);
  if (tokens.length === 0) return undefined;

  const op = tokens[0];
  const args = tokens.slice(1).map(token => numberOrGuide(token, guides));
  const a = args[0];
  const b = args[1];
  const c = args[2];
  let result: number | undefined;

  switch (op) {
    case "val":
      result = a;
      break;
    case "+-":
      result = a !== undefined && b !== undefined && c !== undefined ? a + b - c : undefined;
      break;
    case "+/":
      result = a !== undefined && b !== undefined && c !== undefined ? (a + b) / c : undefined;
      break;
    case "-":
      result = a !== undefined && b !== undefined ? a - b : undefined;
      break;
    case "+":
      result = a !== undefined && b !== undefined ? a + b : undefined;
      break;
    case "*":
      result = a !== undefined && b !== undefined ? a * b : undefined;
      break;
    case "/":
      result = a !== undefined && b !== undefined && b !== 0 ? a / b : undefined;
      break;
    case "*/":
      result = a !== undefined && b !== undefined && c !== undefined && c !== 0 ? (a * b) / c : undefined;
      break;
    case "min":
      result = a !== undefined && b !== undefined ? Math.min(a, b) : undefined;
      break;
    case "max":
      result = a !== undefined && b !== undefined ? Math.max(a, b) : undefined;
      break;
    case "pin":
      result = a !== undefined && b !== undefined && c !== undefined ? Math.min(Math.max(b, a), c) : undefined;
      break;
    case "abs":
      result = a !== undefined ? Math.abs(a) : undefined;
      break;
    case "sqrt":
      result = a !== undefined && a >= 0 ? Math.sqrt(a) : undefined;
      break;
    case "mod":
      result = a !== undefined && b !== undefined && c !== undefined ? Math.sqrt(a * a + b * b + c * c) : undefined;
      break;
    case "sin":
      result = a !== undefined && b !== undefined ? a * Math.sin(angleRadians(b)) : undefined;
      break;
    case "cos":
      result = a !== undefined && b !== undefined ? a * Math.cos(angleRadians(b)) : undefined;
      break;
    case "tan":
      result = a !== undefined && b !== undefined ? a * Math.tan(angleRadians(b)) : undefined;
      break;
    case "at2":
      result = a !== undefined && b !== undefined ? (Math.atan2(a, b) * 180 * 60000) / Math.PI : undefined;
      break;
    case "cat2":
      result = a !== undefined && b !== undefined && c !== undefined
        ? a * Math.cos(Math.atan2(b, c))
        : undefined;
      break;
    case "sat2":
      result = a !== undefined && b !== undefined && c !== undefined
        ? a * Math.sin(Math.atan2(b, c))
        : undefined;
      break;
    default:
      result = tokens.length === 1 ? numberOrGuide(op, guides) : undefined;
  }

  return result !== undefined && Number.isFinite(result) ? result : undefined;
}

function seedGuides(width: number, height: number): GuideTable {
  return {
    w: width,
    h: height,
    l: 0,
    t: 0,
    r: width,
    b: height,
    hc: width / 2,
    vc: height / 2,
    wd2: width / 2,
    hd2: height / 2,
    wd3: width / 3,
    hd3: height / 3,
    wd4: width / 4,
    hd4: height / 4,
    wd5: width / 5,
    hd5: height / 5,
    cd2: 10800000,
    cd4: 5400000,
    "3cd4": 16200000,
    "0": 0,
    "1": 1
  };
}

function readGuides(custGeom: Element, width: number, height: number): GuideTable {
  const guides = seedGuides(width, height);
  const avLst = getDirectChild(custGeom, "avLst");
  const gdLst = getDirectChild(custGeom, "gdLst");
  for (const guide of [
    ...(avLst ? getDirectChildren(avLst, "gd") : []),
    ...(gdLst ? getDirectChildren(gdLst, "gd") : [])
  ]) {
    const name = guide.getAttribute("name");
    const value = evaluateFormula(guide.getAttribute("fmla"), guides);
    if (name && value !== undefined) guides[name] = value;
  }
  return guides;
}

function point(command: Element, guides: GuideTable): { x: number; y: number } | undefined {
  const x = numberOrGuide(command.getAttribute("x"), guides);
  const y = numberOrGuide(command.getAttribute("y"), guides);
  return x !== undefined && y !== undefined ? { x, y } : undefined;
}

function parsePathCommand(node: Element, guides: GuideTable): CustomGeometryCommand | undefined {
  const name = node.localName || node.nodeName.split(":").pop();
  if (name === "close") return { type: "close" };

  if (name === "moveTo" || name === "lnTo") {
    const pt = getDirectChild(node, "pt");
    const value = pt ? point(pt, guides) : undefined;
    if (!value) return undefined;
    return { type: name === "moveTo" ? "moveTo" : "lineTo", ...value };
  }

  if (name === "quadBezTo") {
    const points = getDirectChildren(node, "pt").map(item => point(item, guides)).filter(Boolean) as Array<{ x: number; y: number }>;
    if (points.length < 2) return undefined;
    return {
      type: "quadBezier",
      controlX: points[0].x,
      controlY: points[0].y,
      x: points[1].x,
      y: points[1].y
    };
  }

  if (name === "cubicBezTo") {
    const points = getDirectChildren(node, "pt").map(item => point(item, guides)).filter(Boolean) as Array<{ x: number; y: number }>;
    if (points.length < 3) return undefined;
    return {
      type: "cubicBezier",
      control1X: points[0].x,
      control1Y: points[0].y,
      control2X: points[1].x,
      control2Y: points[1].y,
      x: points[2].x,
      y: points[2].y
    };
  }

  if (name === "arcTo") {
    const widthRadius = numberOrGuide(node.getAttribute("wR"), guides);
    const heightRadius = numberOrGuide(node.getAttribute("hR"), guides);
    const startAngle = numberOrGuide(node.getAttribute("stAng"), guides);
    const sweepAngle = numberOrGuide(node.getAttribute("swAng"), guides);
    if ([widthRadius, heightRadius, startAngle, sweepAngle].some(value => value === undefined)) return undefined;
    return {
      type: "arc",
      widthRadius: widthRadius as number,
      heightRadius: heightRadius as number,
      startAngle: startAngle as number,
      sweepAngle: sweepAngle as number
    };
  }

  return undefined;
}

export function parseCustomGeometry(custGeom: Element): CustomGeometry | undefined {
  const pathLst = getDirectChild(custGeom, "pathLst");
  const pathNodes = pathLst ? getDirectChildren(pathLst, "path") : [];
  if (pathNodes.length === 0) return undefined;

  const firstWidth = numberOrGuide(pathNodes[0].getAttribute("w"), {}) || 1;
  const firstHeight = numberOrGuide(pathNodes[0].getAttribute("h"), {}) || 1;
  const guides = readGuides(custGeom, firstWidth, firstHeight);
  const paths: CustomGeometryPath[] = [];

  for (const pathNode of pathNodes) {
    const width = numberOrGuide(pathNode.getAttribute("w"), guides) || firstWidth;
    const height = numberOrGuide(pathNode.getAttribute("h"), guides) || firstHeight;
    const commands = getDirectChildren(pathNode, "moveTo", "lnTo", "quadBezTo", "cubicBezTo", "arcTo", "close")
      .map(node => parsePathCommand(node, guides))
      .filter(Boolean) as CustomGeometryCommand[];
    const rawFill = pathNode.getAttribute("fill");
    const fill = rawFill === "none" || rawFill === "norm" || rawFill === "lighten" || rawFill === "lightenLess" || rawFill === "darken" || rawFill === "darkenLess"
      ? rawFill
      : undefined;
    const strokeValue = pathNode.getAttribute("stroke");
    const stroke = strokeValue === null
      ? undefined
      : strokeValue === "1" || strokeValue === "true";
    if (commands.length > 0) paths.push({ width, height, commands, fill, stroke });
  }

  return paths.length > 0 ? { paths } : undefined;
}
