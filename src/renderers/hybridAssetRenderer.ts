import type { ExtendedShapeOptions } from '../core/ExtendedPptxJS';
import {
  parseBorder,
  parseBorderRadius,
  parseCssColor,
  parseFilterBlur,
  parseLinearGradient,
  splitTopLevel,
  type CssStyleInput,
  type ParsedColor,
  type RectPx,
} from '../parsers/cssParser';

export const DEFAULT_RASTER_DPI_SCALE = 3; // 3x DPI => 1920x1080 becomes 5760x3240

export interface RenderedAsset {
  dataUrl: string;
  xPx: number;
  yPx: number;
  wPx: number;
  hPx: number;
  dpiScale: number;
}

export type SvgRenderResult =
  | { kind: 'vector'; shapes: ExtendedShapeOptions[] }
  | { kind: 'raster'; asset: RenderedAsset };

interface RadialStop {
  color: ParsedColor;
  pos: number; // 0..1
}

interface ParsedRadialGradient {
  cxRatio: number;
  cyRatio: number;
  stops: RadialStop[];
}

function getStyleValue(style: CssStyleInput, camel: string, kebab: string): string {
  const rec = style as Record<string, unknown>;
  const direct = rec[camel] ?? rec[kebab];
  if (typeof direct === 'string') return direct.trim();
  if (typeof (style as CSSStyleDeclaration).getPropertyValue === 'function') {
    return (style as CSSStyleDeclaration).getPropertyValue(kebab).trim();
  }
  return '';
}

export function getBackdropFilterValue(style: CssStyleInput): string {
  return (
    getStyleValue(style, 'backdropFilter', 'backdrop-filter') ||
    getStyleValue(style, 'webkitBackdropFilter', '-webkit-backdrop-filter')
  );
}

export function hasBackdropBlur(style: CssStyleInput): boolean {
  const val = getBackdropFilterValue(style);
  return Boolean(val && val !== 'none' && /blur\(/i.test(val));
}

export function isComplexGradientOrMesh(style: CssStyleInput): boolean {
  const bgImg = getStyleValue(style, 'backgroundImage', 'background-image');
  const bg = getStyleValue(style, 'background', 'background');
  const combined = `${bgImg} ${bg}`.toLowerCase();

  if (combined.includes('radial-gradient(') || combined.includes('conic-gradient(')) {
    return true;
  }

  // Multiple layered CSS background images separated by top-level commas
  if (bgImg && bgImg !== 'none' && splitTopLevel(bgImg, ',').length > 1) {
    return true;
  }

  return false;
}

export function parseBackdropFilter(style: CssStyleInput): { blurPx: number; saturatePct: number } {
  const raw = getBackdropFilterValue(style);
  const blurMatch = /blur\(\s*([\d.]+)px\s*\)/i.exec(raw);
  const satMatch = /saturate\(\s*([\d.]+)(%?)\s*\)/i.exec(raw);

  const blurPx = blurMatch ? parseFloat(blurMatch[1]) : 0;
  let saturatePct = 100;
  if (satMatch) {
    const val = parseFloat(satMatch[1]);
    saturatePct = satMatch[2] === '%' ? val : val * 100;
  }
  return { blurPx, saturatePct };
}

function extractBalanced(str: string, fnName: string): string | undefined {
  const idx = str.toLowerCase().indexOf(`${fnName.toLowerCase()}(`);
  if (idx === -1) return undefined;
  const openIdx = str.indexOf('(', idx);
  if (openIdx === -1) return undefined;

  let depth = 1;
  for (let i = openIdx + 1; i < str.length; i++) {
    if (str[i] === '(') depth++;
    else if (str[i] === ')') {
      depth--;
      if (depth === 0) return str.slice(openIdx + 1, i);
    }
  }
  return undefined;
}

export function parseRadialGradient(cssValue: string): ParsedRadialGradient | undefined {
  const inner = extractBalanced(cssValue, 'radial-gradient');
  if (!inner) return undefined;

  const parts = splitTopLevel(inner, ',');
  if (parts.length < 2) return undefined;

  let cxRatio = 0.5;
  let cyRatio = 0.5;
  let stopStart = 0;

  const first = parts[0].toLowerCase();
  const isConfig =
    first.includes('circle') ||
    first.includes('ellipse') ||
    first.includes('at ') ||
    first.includes('closest-') ||
    first.includes('farthest-');

  if (isConfig) {
    stopStart = 1;
    const atMatch = /at\s+([^,]+)/i.exec(first);
    if (atMatch) {
      const tokens = atMatch[1].trim().split(/\s+/);
      const parsePos = (tok: string, isX: boolean): number => {
        if (tok === 'center') return 0.5;
        if (tok === 'left') return 0;
        if (tok === 'right') return 1;
        if (tok === 'top') return 0;
        if (tok === 'bottom') return 1;
        if (tok.endsWith('%')) return Math.max(0, Math.min(1, parseFloat(tok) / 100));
        return isX ? 0.5 : 0.5;
      };
      if (tokens[0]) cxRatio = parsePos(tokens[0], true);
      if (tokens[1]) cyRatio = parsePos(tokens[1], false);
    }
  }

  const rawStops = parts.slice(stopStart);
  if (rawStops.length < 2) return undefined;

  const stops: RadialStop[] = [];
  rawStops.forEach((raw, idx) => {
    const match = /^(.*?)(\d+(?:\.\d+)?)%\s*$/.exec(raw.trim());
    let colorStr = raw.trim();
    let pos = idx / Math.max(1, rawStops.length - 1);
    if (match) {
      colorStr = match[1].trim();
      pos = Math.max(0, Math.min(1, parseFloat(match[2]) / 100));
    }
    const parsedColor = parseCssColor(colorStr);
    if (parsedColor) {
      stops.push({ color: parsedColor, pos });
    }
  });

  if (stops.length < 2) return undefined;
  return { cxRatio, cyRatio, stops };
}

function colorToRgbaCss(color: ParsedColor, extraOpacity = 1): string {
  const hex = color.hex;
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  const a = Math.max(0, Math.min(1, color.opacity * extraOpacity));
  return `rgba(${r}, ${g}, ${b}, ${a})`;
}

function traceRoundedRect(
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  w: number,
  h: number,
  r: number
): void {
  const radius = Math.max(0, Math.min(r, Math.min(w, h) / 2));
  ctx.beginPath();
  if (typeof ctx.roundRect === 'function') {
    ctx.roundRect(x, y, w, h, radius);
  } else {
    ctx.moveTo(x + radius, y);
    ctx.lineTo(x + w - radius, y);
    ctx.quadraticCurveTo(x + w, y, x + w, y + radius);
    ctx.lineTo(x + w, y + h - radius);
    ctx.quadraticCurveTo(x + w, y + h, x + w - radius, y + h);
    ctx.lineTo(x + radius, y + h);
    ctx.quadraticCurveTo(x, y + h, x, y + h - radius);
    ctx.lineTo(x, y + radius);
    ctx.quadraticCurveTo(x, y, x + radius, y);
    ctx.closePath();
  }
}

/**
 * Paints a CSS layer (solid, linear-gradient, radial-gradient, blur filter) onto a Canvas 2D context.
 * Useful both for standalone 3x DPI rasterization and for building a slide backdrop canvas
 * so frosted-glass cards can accurately blur layers underneath them.
 */
export function paintCssLayerToCanvas(
  ctx: CanvasRenderingContext2D,
  style: CssStyleInput,
  rect: RectPx
): void {
  const { xPx, yPx, wPx, hPx } = rect;
  if (wPx <= 0 || hPx <= 0) return;

  const opacityRaw = getStyleValue(style, 'opacity', 'opacity');
  const elemOpacity = opacityRaw !== '' && !Number.isNaN(parseFloat(opacityRaw))
    ? Math.max(0, Math.min(1, parseFloat(opacityRaw)))
    : 1;

  const borderRadius = getStyleValue(style, 'borderRadius', 'border-radius');
  const radiusPx = borderRadius === '50%'
    ? Math.min(wPx, hPx) / 2
    : parseBorderRadius(borderRadius, wPx, hPx) ?? 0;

  const filterRaw = getStyleValue(style, 'filter', 'filter');
  const blurPt = parseFilterBlur(filterRaw, 0.5);
  const blurPx = blurPt !== undefined ? blurPt * 2 : 0;

  ctx.save();
  if (blurPx > 0) {
    ctx.filter = `blur(${blurPx}px)`;
  }

  traceRoundedRect(ctx, xPx, yPx, wPx, hPx, radiusPx);
  ctx.clip();

  const bgColorStr = getStyleValue(style, 'backgroundColor', 'background-color');
  const bgColor = parseCssColor(bgColorStr);
  if (bgColor && bgColor.opacity > 0) {
    ctx.fillStyle = colorToRgbaCss(bgColor, elemOpacity);
    ctx.fillRect(xPx, yPx, wPx, hPx);
  }

  const bgImg = getStyleValue(style, 'backgroundImage', 'background-image');
  const bg = getStyleValue(style, 'background', 'background');
  const gradSource = bgImg && bgImg !== 'none' ? bgImg : bg;

  if (gradSource && gradSource !== 'none') {
    const layers = splitTopLevel(gradSource, ',');
    // CSS renders background layers top-to-bottom (first is topmost), so paint in reverse
    for (let i = layers.length - 1; i >= 0; i--) {
      const layer = layers[i];
      if (/radial-gradient\(/i.test(layer)) {
        const radial = parseRadialGradient(layer);
        if (radial) {
          const cx = xPx + wPx * radial.cxRatio;
          const cy = yPx + hPx * radial.cyRatio;
          const rMax = Math.hypot(Math.max(cx - xPx, xPx + wPx - cx), Math.max(cy - yPx, yPx + hPx - cy));
          const grad = ctx.createRadialGradient(cx, cy, 0, cx, cy, Math.max(1, rMax));
          for (const stop of radial.stops) {
            grad.addColorStop(stop.pos, colorToRgbaCss(stop.color, elemOpacity));
          }
          ctx.fillStyle = grad;
          ctx.fillRect(xPx, yPx, wPx, hPx);
        }
      } else if (/linear-gradient\(/i.test(layer)) {
        const linear = parseLinearGradient(layer);
        if (linear) {
          const rad = ((linear.angleDeg - 90) * Math.PI) / 180;
          const halfDiag = Math.hypot(wPx, hPx) / 2;
          const cx = xPx + wPx / 2;
          const cy = yPx + hPx / 2;
          const x0 = cx - Math.cos(rad) * halfDiag;
          const y0 = cy - Math.sin(rad) * halfDiag;
          const x1 = cx + Math.cos(rad) * halfDiag;
          const y1 = cy + Math.sin(rad) * halfDiag;
          const grad = ctx.createLinearGradient(x0, y0, x1, y1);
          for (const stop of linear.stops) {
            const c: ParsedColor = { hex: stop.color, opacity: stop.opacity ?? 1 };
            grad.addColorStop(stop.position / 100, colorToRgbaCss(c, elemOpacity));
          }
          ctx.fillStyle = grad;
          ctx.fillRect(xPx, yPx, wPx, hPx);
        }
      }
    }
  }

  ctx.restore();
}

/**
 * Renders a CSS `backdrop-filter: blur()` frosted-glass card onto an offscreen
 * HTML5 Canvas at 3x DPI (5760x3240 slide equivalent) and returns a transparent PNG.
 */
export async function renderFrostedGlassElement(
  style: CssStyleInput,
  rect: RectPx,
  backdropCanvas?: HTMLCanvasElement,
  dpiScale = DEFAULT_RASTER_DPI_SCALE
): Promise<RenderedAsset> {
  const { xPx, yPx, wPx, hPx } = rect;
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.round(wPx * dpiScale));
  canvas.height = Math.max(1, Math.round(hPx * dpiScale));

  const ctx = canvas.getContext('2d');
  if (!ctx) {
    throw new Error('Failed to acquire 2D canvas context for frosted-glass rasterization.');
  }

  ctx.scale(dpiScale, dpiScale);

  const borderRadius = getStyleValue(style, 'borderRadius', 'border-radius');
  const radiusPx = borderRadius === '50%'
    ? Math.min(wPx, hPx) / 2
    : parseBorderRadius(borderRadius, wPx, hPx) ?? 0;

  const { blurPx, saturatePct } = parseBackdropFilter(style);

  ctx.save();
  traceRoundedRect(ctx, 0, 0, wPx, hPx, radiusPx);
  ctx.clip();

  // 1. Sample & blur backdrop if a slide backdrop canvas was captured
  if (backdropCanvas && backdropCanvas.width > 0 && backdropCanvas.height > 0) {
    ctx.save();
    const filterParts: string[] = [];
    if (blurPx > 0) filterParts.push(`blur(${blurPx}px)`);
    if (saturatePct !== 100) filterParts.push(`saturate(${saturatePct}%)`);
    if (filterParts.length > 0) {
      ctx.filter = filterParts.join(' ');
    }

    const scaleX = backdropCanvas.width / 1920;
    const scaleY = backdropCanvas.height / 1080;
    const pad = Math.ceil(blurPx * 1.5);
    const sx = Math.max(0, (xPx - pad) * scaleX);
    const sy = Math.max(0, (yPx - pad) * scaleY);
    const sw = Math.min(backdropCanvas.width - sx, (wPx + pad * 2) * scaleX);
    const sh = Math.min(backdropCanvas.height - sy, (hPx + pad * 2) * scaleY);
    const dx = sx / scaleX - xPx;
    const dy = sy / scaleY - yPx;
    const dw = sw / scaleX;
    const dh = sh / scaleY;

    if (sw > 0 && sh > 0) {
      ctx.drawImage(backdropCanvas, sx, sy, sw, sh, dx, dy, dw, dh);
    }
    ctx.restore();
  }

  // 2. Apply card surface fill (translucent glass tint / gradient)
  paintCssLayerToCanvas(ctx, style, { xPx: 0, yPx: 0, wPx, hPx });

  // 3. Subtle top-lit frosted specular sheen
  const sheen = ctx.createLinearGradient(0, 0, 0, hPx);
  sheen.addColorStop(0, 'rgba(255, 255, 255, 0.14)');
  sheen.addColorStop(0.25, 'rgba(255, 255, 255, 0.03)');
  sheen.addColorStop(1, 'rgba(255, 255, 255, 0.0)');
  ctx.fillStyle = sheen;
  ctx.fillRect(0, 0, wPx, hPx);

  ctx.restore();

  // 4. Crisp vector border / top-lit bevel stroke
  const borderRaw = getStyleValue(style, 'border', 'border');
  const parsedBorder = parseBorder(borderRaw, 1); // 1x px here since ctx is scaled by dpiScale
  const topWidthStr = getStyleValue(style, 'borderTopWidth', 'border-top-width');
  const topColorStr = getStyleValue(style, 'borderTopColor', 'border-top-color');

  if (parsedBorder) {
    ctx.save();
    traceRoundedRect(ctx, 0.5, 0.5, Math.max(0, wPx - 1), Math.max(0, hPx - 1), radiusPx);
    ctx.lineWidth = parsedBorder.widthPt;
    ctx.strokeStyle = colorToRgbaCss({
      hex: parsedBorder.color,
      opacity: parsedBorder.opacity ?? 1,
    });
    ctx.stroke();
    ctx.restore();
  } else if (topWidthStr && topColorStr) {
    const topColor = parseCssColor(topColorStr);
    const topWidth = parseFloat(topWidthStr);
    if (topColor && !Number.isNaN(topWidth) && topWidth > 0 && topColor.opacity > 0) {
      ctx.save();
      traceRoundedRect(ctx, 0.5, 0.5, Math.max(0, wPx - 1), Math.max(0, hPx - 1), radiusPx);
      ctx.lineWidth = topWidth;
      ctx.strokeStyle = colorToRgbaCss(topColor);
      ctx.stroke();
      ctx.restore();
    }
  }

  return {
    dataUrl: canvas.toDataURL('image/png'),
    xPx,
    yPx,
    wPx,
    hPx,
    dpiScale,
  };
}

/**
 * Renders multi-layer mesh blurs / radial-gradient / CSS blur elements onto a
 * 3x DPI offscreen HTML5 Canvas and returns a transparent PNG.
 */
export async function renderMeshBlurElement(
  style: CssStyleInput,
  rect: RectPx,
  dpiScale = DEFAULT_RASTER_DPI_SCALE
): Promise<RenderedAsset> {
  const filterRaw = getStyleValue(style, 'filter', 'filter');
  const blurPt = parseFilterBlur(filterRaw, 0.5);
  const blurPx = blurPt !== undefined ? blurPt * 2 : 0;
  const padPx = blurPx > 0 ? Math.ceil(blurPx * 2) : 0;

  const outX = rect.xPx - padPx;
  const outY = rect.yPx - padPx;
  const outW = rect.wPx + padPx * 2;
  const outH = rect.hPx + padPx * 2;

  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.round(outW * dpiScale));
  canvas.height = Math.max(1, Math.round(outH * dpiScale));

  const ctx = canvas.getContext('2d');
  if (!ctx) {
    throw new Error('Failed to acquire 2D canvas context for mesh blur rasterization.');
  }

  ctx.scale(dpiScale, dpiScale);
  paintCssLayerToCanvas(ctx, style, {
    xPx: padPx,
    yPx: padPx,
    wPx: rect.wPx,
    hPx: rect.hPx,
  });

  return {
    dataUrl: canvas.toDataURL('image/png'),
    xPx: outX,
    yPx: outY,
    wPx: outW,
    hPx: outH,
    dpiScale,
  };
}

/**
 * Determines whether an SVG requires 3x DPI raster fallback vs native <a:custGeom> conversion.
 */
export function isComplexSvg(svgEl: SVGSVGElement): boolean {
  const complexTags = 'filter, linearGradient, radialGradient, mask, clipPath, pattern, text, image, use';
  if (svgEl.querySelector(complexTags)) return true;

  const descendants = svgEl.querySelectorAll('*');
  for (const el of Array.from(descendants)) {
    const attrs = ['filter', 'mask', 'clip-path', 'stroke-dasharray', 'transform'];
    for (const attr of attrs) {
      const val = el.getAttribute(attr);
      if (val && val !== 'none') return true;
    }
    const fill = el.getAttribute('fill') ?? '';
    const stroke = el.getAttribute('stroke') ?? '';
    if (fill.includes('url(') || stroke.includes('url(')) return true;
  }

  return false;
}

function parseSvgViewBox(svgEl: SVGSVGElement, rect: RectPx): { minX: number; minY: number; vbW: number; vbH: number } {
  const vb = svgEl.getAttribute('viewBox');
  if (vb) {
    const nums = vb.trim().split(/[\s,]+/).map(Number);
    if (nums.length === 4 && nums.every((n) => !Number.isNaN(n)) && nums[2] > 0 && nums[3] > 0) {
      return { minX: nums[0], minY: nums[1], vbW: nums[2], vbH: nums[3] };
    }
  }
  return { minX: 0, minY: 0, vbW: Math.max(1, rect.wPx), vbH: Math.max(1, rect.hPx) };
}

function buildPolygonCustGeomXml(
  points: Array<[number, number]>,
  vbW: number,
  vbH: number,
  closed: boolean
): string {
  const coordScale = 10000;
  const w = Math.max(1, Math.round(vbW * coordScale));
  const h = Math.max(1, Math.round(vbH * coordScale));

  const cmds = points
    .map(([px, py], idx) => {
      const x = Math.round(px * coordScale);
      const y = Math.round(py * coordScale);
      return idx === 0 ? `<a:moveTo><a:pt x="${x}" y="${y}"/></a:moveTo>` : `<a:lnTo><a:pt x="${x}" y="${y}"/></a:lnTo>`;
    })
    .join('');

  const closeTag = closed ? '<a:close/>' : '';
  return `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="${w}" h="${h}">${cmds}${closeTag}</a:path></a:pathLst></a:custGeom>`;
}

function svgPathToCustGeomXml(d: string, minX: number, minY: number, vbW: number, vbH: number): string | undefined {
  const tokens = d.match(/[a-zA-Z]|-?\d*\.?\d+(?:e[-+]?\d+)?/g);
  if (!tokens || tokens.length === 0) return undefined;

  const coordScale = 10000;
  const w = Math.max(1, Math.round(vbW * coordScale));
  const h = Math.max(1, Math.round(vbH * coordScale));

  const toX = (val: number): number => Math.round((val - minX) * coordScale);
  const toY = (val: number): number => Math.round((val - minY) * coordScale);

  let xml = '';
  let idx = 0;
  let curX = 0;
  let curY = 0;
  let startX = 0;
  let startY = 0;
  let cmd = '';

  const nextNum = (): number => parseFloat(tokens[idx++]);

  while (idx < tokens.length) {
    const tok = tokens[idx];
    if (/^[a-zA-Z]$/.test(tok)) {
      cmd = tok;
      idx++;
      if (cmd === 'Z' || cmd === 'z') {
        xml += '<a:close/>';
        curX = startX;
        curY = startY;
      }
      continue;
    }

    switch (cmd) {
      case 'M':
      case 'm': {
        const x = nextNum() + (cmd === 'm' ? curX : 0);
        const y = nextNum() + (cmd === 'm' ? curY : 0);
        curX = x;
        curY = y;
        startX = x;
        startY = y;
        xml += `<a:moveTo><a:pt x="${toX(x)}" y="${toY(y)}"/></a:moveTo>`;
        cmd = cmd === 'm' ? 'l' : 'L';
        break;
      }
      case 'L':
      case 'l': {
        const x = nextNum() + (cmd === 'l' ? curX : 0);
        const y = nextNum() + (cmd === 'l' ? curY : 0);
        curX = x;
        curY = y;
        xml += `<a:lnTo><a:pt x="${toX(x)}" y="${toY(y)}"/></a:lnTo>`;
        break;
      }
      case 'H':
      case 'h': {
        const x = nextNum() + (cmd === 'h' ? curX : 0);
        curX = x;
        xml += `<a:lnTo><a:pt x="${toX(curX)}" y="${toY(curY)}"/></a:lnTo>`;
        break;
      }
      case 'V':
      case 'v': {
        const y = nextNum() + (cmd === 'v' ? curY : 0);
        curY = y;
        xml += `<a:lnTo><a:pt x="${toX(curX)}" y="${toY(curY)}"/></a:lnTo>`;
        break;
      }
      case 'C':
      case 'c': {
        const relX = cmd === 'c' ? curX : 0;
        const relY = cmd === 'c' ? curY : 0;
        const x1 = nextNum() + relX;
        const y1 = nextNum() + relY;
        const x2 = nextNum() + relX;
        const y2 = nextNum() + relY;
        const x = nextNum() + relX;
        const y = nextNum() + relY;
        curX = x;
        curY = y;
        xml +=
          `<a:cubicBezTo>` +
          `<a:pt x="${toX(x1)}" y="${toY(y1)}"/>` +
          `<a:pt x="${toX(x2)}" y="${toY(y2)}"/>` +
          `<a:pt x="${toX(x)}" y="${toY(y)}"/>` +
          `</a:cubicBezTo>`;
        break;
      }
      case 'Q':
      case 'q': {
        const relX = cmd === 'q' ? curX : 0;
        const relY = cmd === 'q' ? curY : 0;
        const x1 = nextNum() + relX;
        const y1 = nextNum() + relY;
        const x = nextNum() + relX;
        const y = nextNum() + relY;
        curX = x;
        curY = y;
        xml +=
          `<a:quadBezTo>` +
          `<a:pt x="${toX(x1)}" y="${toY(y1)}"/>` +
          `<a:pt x="${toX(x)}" y="${toY(y)}"/>` +
          `</a:quadBezTo>`;
        break;
      }
      default:
        // Unsupported command (e.g., elliptical arc A/a) -> fallback to raster
        return undefined;
    }
  }

  if (!xml) return undefined;
  return `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="${w}" h="${h}">${xml}</a:path></a:pathLst></a:custGeom>`;
}

/**
 * Converts simple geometric SVG primitives (`rect`, `circle`, `ellipse`, `line`, `polygon`,
 * `polyline`, `path`) into native editable PowerPoint shapes (`ExtendedShapeOptions[]`).
 * Returns `null` if any element cannot be represented cleanly as a native vector.
 */
export function convertSimpleSvgToShapes(
  svgEl: SVGSVGElement,
  rect: RectPx
): ExtendedShapeOptions[] | null {
  if (isComplexSvg(svgEl)) return null;

  const { minX, minY, vbW, vbH } = parseSvgViewBox(svgEl, rect);
  const scaleX = rect.wPx / vbW;
  const scaleY = rect.hPx / vbH;

  const shapes: ExtendedShapeOptions[] = [];
  const children = svgEl.querySelectorAll('rect, circle, ellipse, line, polygon, polyline, path');

  const svgComputedColor =
    typeof window !== 'undefined' ? window.getComputedStyle(svgEl).color : undefined;

  for (const node of Array.from(children)) {
    const tag = node.tagName.toLowerCase();
    let fillAttr = node.getAttribute('fill') ?? svgEl.getAttribute('fill') ?? '#000000';
    let strokeAttr = node.getAttribute('stroke') ?? svgEl.getAttribute('stroke') ?? 'none';
    if (fillAttr.toLowerCase() === 'currentcolor' && svgComputedColor) {
      fillAttr = svgComputedColor;
    }
    if (strokeAttr.toLowerCase() === 'currentcolor' && svgComputedColor) {
      strokeAttr = svgComputedColor;
    }
    const strokeWidthAttr = parseFloat(
      node.getAttribute('stroke-width') ?? svgEl.getAttribute('stroke-width') ?? '1'
    );
    const opacityAttr = parseFloat(node.getAttribute('opacity') ?? '1');
    const fillOpacityAttr = parseFloat(node.getAttribute('fill-opacity') ?? '1');
    const strokeOpacityAttr = parseFloat(node.getAttribute('stroke-opacity') ?? '1');

    const parsedFill = fillAttr !== 'none' ? parseCssColor(fillAttr) : undefined;
    const parsedStroke = strokeAttr !== 'none' ? parseCssColor(strokeAttr) : undefined;

    const baseFillOpacity = parsedFill
      ? Number((parsedFill.opacity * opacityAttr * fillOpacityAttr).toFixed(4))
      : 0;

    const border =
      parsedStroke && parsedStroke.opacity > 0 && strokeWidthAttr > 0
        ? {
            color: parsedStroke.hex,
            widthPt: Number((strokeWidthAttr * Math.min(scaleX, scaleY) * 0.5).toFixed(2)),
            opacity: Number((parsedStroke.opacity * opacityAttr * strokeOpacityAttr).toFixed(4)),
          }
        : undefined;

    if (tag === 'rect') {
      const x = parseFloat(node.getAttribute('x') ?? '0');
      const y = parseFloat(node.getAttribute('y') ?? '0');
      const w = parseFloat(node.getAttribute('width') ?? '0');
      const h = parseFloat(node.getAttribute('height') ?? '0');
      const rx = parseFloat(node.getAttribute('rx') ?? node.getAttribute('ry') ?? '0');
      if (w <= 0 || h <= 0) continue;

      shapes.push({
        xPx: rect.xPx + (x - minX) * scaleX,
        yPx: rect.yPx + (y - minY) * scaleY,
        wPx: w * scaleX,
        hPx: h * scaleY,
        shapeType: rx > 0 ? 'roundRect' : 'rect',
        cornerRadiusPx: rx > 0 ? rx * Math.min(scaleX, scaleY) : undefined,
        fillColor: parsedFill?.hex ?? 'FFFFFF',
        fillOpacity: parsedFill ? baseFillOpacity : 0,
        border,
      });
    } else if (tag === 'circle' || tag === 'ellipse') {
      const cx = parseFloat(node.getAttribute('cx') ?? '0');
      const cy = parseFloat(node.getAttribute('cy') ?? '0');
      const rx = parseFloat(node.getAttribute(tag === 'circle' ? 'r' : 'rx') ?? '0');
      const ry = parseFloat(node.getAttribute(tag === 'circle' ? 'r' : 'ry') ?? '0');
      if (rx <= 0 || ry <= 0) continue;

      shapes.push({
        xPx: rect.xPx + (cx - rx - minX) * scaleX,
        yPx: rect.yPx + (cy - ry - minY) * scaleY,
        wPx: rx * 2 * scaleX,
        hPx: ry * 2 * scaleY,
        shapeType: 'ellipse',
        fillColor: parsedFill?.hex ?? 'FFFFFF',
        fillOpacity: parsedFill ? baseFillOpacity : 0,
        border,
      });
    } else if (tag === 'polygon' || tag === 'polyline' || tag === 'line') {
      const pts: Array<[number, number]> = [];
      if (tag === 'line') {
        pts.push([
          parseFloat(node.getAttribute('x1') ?? '0') - minX,
          parseFloat(node.getAttribute('y1') ?? '0') - minY,
        ]);
        pts.push([
          parseFloat(node.getAttribute('x2') ?? '0') - minX,
          parseFloat(node.getAttribute('y2') ?? '0') - minY,
        ]);
      } else {
        const rawPts = (node.getAttribute('points') ?? '').trim().split(/[\s,]+/).map(Number);
        for (let i = 0; i + 1 < rawPts.length; i += 2) {
          pts.push([rawPts[i] - minX, rawPts[i + 1] - minY]);
        }
      }
      if (pts.length < 2) continue;

      const customGeometryXml = buildPolygonCustGeomXml(pts, vbW, vbH, tag === 'polygon');
      shapes.push({
        xPx: rect.xPx,
        yPx: rect.yPx,
        wPx: rect.wPx,
        hPx: rect.hPx,
        shapeType: 'rect',
        fillColor: tag === 'polygon' && parsedFill ? parsedFill.hex : 'FFFFFF',
        fillOpacity: tag === 'polygon' && parsedFill ? baseFillOpacity : 0,
        border,
        customGeometryXml,
      });
    } else if (tag === 'path') {
      const d = node.getAttribute('d') ?? '';
      const customGeometryXml = svgPathToCustGeomXml(d, minX, minY, vbW, vbH);
      if (!customGeometryXml) return null;

      shapes.push({
        xPx: rect.xPx,
        yPx: rect.yPx,
        wPx: rect.wPx,
        hPx: rect.hPx,
        shapeType: 'rect',
        fillColor: parsedFill?.hex ?? 'FFFFFF',
        fillOpacity: parsedFill ? baseFillOpacity : 0,
        border,
        customGeometryXml,
      });
    }
  }

  return shapes.length > 0 ? shapes : null;
}

/**
 * Rasterizes an SVG element onto an offscreen HTML5 Canvas at 3x DPI as a transparent PNG.
 */
export async function rasterizeSvgToPng(
  svgEl: SVGSVGElement,
  rect: RectPx,
  dpiScale = DEFAULT_RASTER_DPI_SCALE
): Promise<RenderedAsset> {
  const clone = svgEl.cloneNode(true) as SVGSVGElement;
  clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
  clone.setAttribute('width', String(rect.wPx));
  clone.setAttribute('height', String(rect.hPx));
  if (typeof window !== 'undefined') {
    const computedColor = window.getComputedStyle(svgEl).color;
    if (computedColor) {
      clone.style.color = computedColor;
    }
  }

  const xml = new XMLSerializer().serializeToString(clone);
  const svgDataUrl = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(xml)}`;

  const img = await new Promise<HTMLImageElement>((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error('Failed to load serialized SVG into Image for rasterization.'));
    image.src = svgDataUrl;
  });

  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.round(rect.wPx * dpiScale));
  canvas.height = Math.max(1, Math.round(rect.hPx * dpiScale));

  const ctx = canvas.getContext('2d');
  if (!ctx) {
    throw new Error('Failed to acquire 2D canvas context for SVG rasterization.');
  }
  ctx.scale(dpiScale, dpiScale);
  ctx.drawImage(img, 0, 0, rect.wPx, rect.hPx);

  return {
    dataUrl: canvas.toDataURL('image/png'),
    xPx: rect.xPx,
    yPx: rect.yPx,
    wPx: rect.wPx,
    hPx: rect.hPx,
    dpiScale,
  };
}

/**
 * Handles an SVG element: converts simple geometric SVGs to native `<a:custGeom>` /
 * geometric PowerPoint vectors, and falls back to 3x DPI PNG rasterization only for
 * complex SVGs (filters, gradients, dashed arcs, masks, etc.).
 */
export async function renderSvgElement(
  svgEl: SVGSVGElement,
  rect: RectPx,
  dpiScale = DEFAULT_RASTER_DPI_SCALE
): Promise<SvgRenderResult> {
  const vectorShapes = convertSimpleSvgToShapes(svgEl, rect);
  if (vectorShapes) {
    return { kind: 'vector', shapes: vectorShapes };
  }
  const asset = await rasterizeSvgToPng(svgEl, rect, dpiScale);
  return { kind: 'raster', asset };
}
