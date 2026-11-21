import type {
  ExtendedGradient,
  ExtendedShadow,
  ExtendedShapeOptions,
  GradientStop,
  PptxShapeName,
} from '../core/ExtendedPptxJS';

export interface ParsedColor {
  hex: string;     // 6-char uppercase hex without '#'
  opacity: number; // 0 to 1
}

export interface RectPx {
  xPx: number;
  yPx: number;
  wPx: number;
  hPx: number;
}

export type CssStyleInput =
  | Partial<Record<string, string | undefined>>
  | CSSStyleDeclaration;

const DEFAULT_PX_TO_PT = 0.5; // 1920px -> 13.333in (960pt) => 0.5 pt/px

const NAMED_COLORS: Record<string, string> = {
  black: '000000',
  white: 'FFFFFF',
  red: 'FF0000',
  green: '008000',
  blue: '0000FF',
  yellow: 'FFFF00',
  cyan: '00FFFF',
  aqua: '00FFFF',
  magenta: 'FF00FF',
  fuchsia: 'FF00FF',
  silver: 'C0C0C0',
  gray: '808080',
  grey: '808080',
  maroon: '800000',
  olive: '808000',
  lime: '00FF00',
  teal: '008080',
  navy: '000080',
  purple: '800080',
  orange: 'FFA500',
  pink: 'FFC0CB',
  rebeccapurple: '663399',
  indigo: '4B0082',
  violet: 'EE82EE',
  coral: 'FF7F50',
  gold: 'FFD700',
  crimson: 'DC143C',
  turquoise: '40E0D0',
  slategray: '708090',
  slategrey: '708090',
};

function clamp(val: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, val));
}

function toHexByte(n: number): string {
  return clamp(Math.round(n), 0, 255).toString(16).padStart(2, '0').toUpperCase();
}

function roundOpacity(opacity: number): number {
  return Number(clamp(opacity, 0, 1).toFixed(4));
}

function hslToRgb(hDeg: number, sPct: number, lPct: number): [number, number, number] {
  const h = ((hDeg % 360) + 360) % 360;
  const s = clamp(sPct, 0, 100) / 100;
  const l = clamp(lPct, 0, 100) / 100;

  const c = (1 - Math.abs(2 * l - 1)) * s;
  const x = c * (1 - Math.abs(((h / 60) % 2) - 1));
  const m = l - c / 2;

  let r1 = 0;
  let g1 = 0;
  let b1 = 0;

  if (h < 60) {
    [r1, g1, b1] = [c, x, 0];
  } else if (h < 120) {
    [r1, g1, b1] = [x, c, 0];
  } else if (h < 180) {
    [r1, g1, b1] = [0, c, x];
  } else if (h < 240) {
    [r1, g1, b1] = [0, x, c];
  } else if (h < 300) {
    [r1, g1, b1] = [x, 0, c];
  } else {
    [r1, g1, b1] = [c, 0, x];
  }

  return [
    Math.round((r1 + m) * 255),
    Math.round((g1 + m) * 255),
    Math.round((b1 + m) * 255),
  ];
}

function parseAlphaToken(token: string): number {
  const trimmed = token.trim();
  if (trimmed.endsWith('%')) {
    return roundOpacity(parseFloat(trimmed) / 100);
  }
  return roundOpacity(parseFloat(trimmed));
}

function parseChannelToken(token: string): number {
  const trimmed = token.trim();
  if (trimmed.endsWith('%')) {
    return clamp((parseFloat(trimmed) / 100) * 255, 0, 255);
  }
  return clamp(parseFloat(trimmed), 0, 255);
}

/**
 * Splits a CSS string on a delimiter only at parenthesis depth 0.
 */
export function splitTopLevel(input: string, delimiter = ','): string[] {
  const result: string[] = [];
  let depth = 0;
  let start = 0;
  let inQuote: string | null = null;

  for (let i = 0; i < input.length; i++) {
    const ch = input[i];
    if (inQuote) {
      if (ch === inQuote && input[i - 1] !== '\\') {
        inQuote = null;
      }
      continue;
    }
    if (ch === '"' || ch === "'") {
      inQuote = ch;
      continue;
    }
    if (ch === '(') {
      depth++;
    } else if (ch === ')') {
      depth = Math.max(0, depth - 1);
    } else if (ch === delimiter && depth === 0) {
      const slice = input.slice(start, i).trim();
      if (slice) result.push(slice);
      start = i + 1;
    }
  }

  const tail = input.slice(start).trim();
  if (tail) result.push(tail);
  return result;
}

/**
 * Parses hex (#RGB, #RGBA, #RRGGBB, #RRGGBBAA), rgb/rgba, hsl/hsla,
 * 'transparent', and standard CSS named colors.
 */
export function parseCssColor(colorStr: string | undefined | null): ParsedColor | undefined {
  if (!colorStr) return undefined;
  const raw = colorStr.trim();
  if (!raw || raw.toLowerCase() === 'none') return undefined;

  const lower = raw.toLowerCase();
  if (lower === 'transparent') {
    return { hex: '000000', opacity: 0 };
  }

  if (NAMED_COLORS[lower]) {
    return { hex: NAMED_COLORS[lower], opacity: 1 };
  }

  // Hex formats
  if (raw.startsWith('#')) {
    const hexBody = raw.slice(1);
    if (/^[0-9a-fA-F]{3}$/.test(hexBody)) {
      const expanded = hexBody
        .split('')
        .map((c) => c + c)
        .join('')
        .toUpperCase();
      return { hex: expanded, opacity: 1 };
    }
    if (/^[0-9a-fA-F]{4}$/.test(hexBody)) {
      const r = hexBody[0] + hexBody[0];
      const g = hexBody[1] + hexBody[1];
      const b = hexBody[2] + hexBody[2];
      const a = hexBody[3] + hexBody[3];
      return {
        hex: `${r}${g}${b}`.toUpperCase(),
        opacity: roundOpacity(parseInt(a, 16) / 255),
      };
    }
    if (/^[0-9a-fA-F]{6}$/.test(hexBody)) {
      return { hex: hexBody.toUpperCase(), opacity: 1 };
    }
    if (/^[0-9a-fA-F]{8}$/.test(hexBody)) {
      return {
        hex: hexBody.slice(0, 6).toUpperCase(),
        opacity: roundOpacity(parseInt(hexBody.slice(6, 8), 16) / 255),
      };
    }
    return undefined;
  }

  // rgb(...) / rgba(...)
  const rgbMatch = /^rgba?\(\s*([^)]+)\s*\)$/i.exec(raw);
  if (rgbMatch) {
    const inner = rgbMatch[1];
    let channels: string[];
    let alphaStr: string | undefined;

    if (inner.includes(',')) {
      const parts = inner.split(',').map((s) => s.trim());
      if (parts.length < 3 || parts.length > 4) return undefined;
      channels = parts.slice(0, 3);
      alphaStr = parts[3];
    } else {
      const slashParts = inner.split('/').map((s) => s.trim());
      channels = slashParts[0].split(/\s+/).filter(Boolean);
      if (channels.length !== 3) return undefined;
      alphaStr = slashParts[1];
    }

    const r = parseChannelToken(channels[0]);
    const g = parseChannelToken(channels[1]);
    const b = parseChannelToken(channels[2]);
    if (Number.isNaN(r) || Number.isNaN(g) || Number.isNaN(b)) return undefined;

    const opacity = alphaStr !== undefined ? parseAlphaToken(alphaStr) : 1;
    if (Number.isNaN(opacity)) return undefined;

    return {
      hex: `${toHexByte(r)}${toHexByte(g)}${toHexByte(b)}`,
      opacity,
    };
  }

  // hsl(...) / hsla(...)
  const hslMatch = /^hsla?\(\s*([^)]+)\s*\)$/i.exec(raw);
  if (hslMatch) {
    const inner = hslMatch[1];
    let parts: string[];
    let alphaStr: string | undefined;

    if (inner.includes(',')) {
      const split = inner.split(',').map((s) => s.trim());
      if (split.length < 3 || split.length > 4) return undefined;
      parts = split.slice(0, 3);
      alphaStr = split[3];
    } else {
      const slashParts = inner.split('/').map((s) => s.trim());
      parts = slashParts[0].split(/\s+/).filter(Boolean);
      if (parts.length !== 3) return undefined;
      alphaStr = slashParts[1];
    }

    const hueDeg = parseAngleToDeg(parts[0]) ?? parseFloat(parts[0]);
    const sat = parseFloat(parts[1]);
    const light = parseFloat(parts[2]);
    if (Number.isNaN(hueDeg) || Number.isNaN(sat) || Number.isNaN(light)) return undefined;

    const [r, g, b] = hslToRgb(hueDeg, sat, light);
    const opacity = alphaStr !== undefined ? parseAlphaToken(alphaStr) : 1;
    if (Number.isNaN(opacity)) return undefined;

    return {
      hex: `${toHexByte(r)}${toHexByte(g)}${toHexByte(b)}`,
      opacity,
    };
  }

  return undefined;
}

function parseAngleToDeg(token: string): number | undefined {
  const lower = token.trim().toLowerCase();
  const numMatch = /^(-?\d*\.?\d+)(deg|rad|grad|turn)?$/.exec(lower);
  if (numMatch) {
    const val = parseFloat(numMatch[1]);
    const unit = numMatch[2];
    if (!unit && val === 0) return 0;
    if (!unit) return undefined;
    switch (unit) {
      case 'deg':
        return ((val % 360) + 360) % 360;
      case 'turn':
        return (((val * 360) % 360) + 360) % 360;
      case 'rad':
        return ((((val * 180) / Math.PI) % 360) + 360) % 360;
      case 'grad':
        return (((val * 0.9) % 360) + 360) % 360;
    }
  }

  if (lower.startsWith('to ')) {
    const dirs = lower
      .slice(3)
      .trim()
      .split(/\s+/)
      .sort()
      .join(' ');
    const dirMap: Record<string, number> = {
      top: 0,
      'right top': 45,
      right: 90,
      'bottom right': 135,
      bottom: 180,
      'bottom left': 225,
      left: 270,
      'left top': 315,
    };
    return dirMap[dirs];
  }

  return undefined;
}

function extractFunctionArgs(cssValue: string, fnName: string): string | undefined {
  const regex = new RegExp(`${fnName}\\s*\\(`, 'i');
  const match = regex.exec(cssValue);
  if (!match) return undefined;

  const startIdx = match.index + match[0].length;
  let depth = 1;
  for (let i = startIdx; i < cssValue.length; i++) {
    const ch = cssValue[i];
    if (ch === '(') depth++;
    else if (ch === ')') {
      depth--;
      if (depth === 0) {
        return cssValue.slice(startIdx, i);
      }
    }
  }
  return undefined;
}

function tokenizeCssStop(stopStr: string): string[] {
  const tokens: string[] = [];
  let current = '';
  let depth = 0;

  for (let i = 0; i < stopStr.length; i++) {
    const ch = stopStr[i];
    if (ch === '(') {
      depth++;
      current += ch;
    } else if (ch === ')') {
      depth = Math.max(0, depth - 1);
      current += ch;
    } else if (/\s/.test(ch) && depth === 0) {
      if (current) {
        tokens.push(current);
        current = '';
      }
    } else {
      current += ch;
    }
  }
  if (current) tokens.push(current);
  return tokens;
}

/**
 * Parses a CSS `linear-gradient(...)` expression into an `ExtendedGradient`.
 */
export function parseLinearGradient(cssValue: string | undefined | null): ExtendedGradient | undefined {
  if (!cssValue) return undefined;
  const inner = extractFunctionArgs(cssValue, 'linear-gradient');
  if (!inner) return undefined;

  const args = splitTopLevel(inner, ',');
  if (args.length < 2) return undefined;

  let angleDeg = 180; // CSS default: "to bottom"
  let stopStartIdx = 0;

  const maybeAngle = parseAngleToDeg(args[0]);
  if (maybeAngle !== undefined) {
    angleDeg = Number(maybeAngle.toFixed(2));
    stopStartIdx = 1;
  }

  const rawStops = args.slice(stopStartIdx);
  if (rawStops.length < 2) return undefined;

  interface TempStop {
    color: ParsedColor;
    pos?: number;
  }
  const expanded: TempStop[] = [];

  for (const rawStop of rawStops) {
    const tokens = tokenizeCssStop(rawStop);
    if (tokens.length === 0) continue;

    const color = parseCssColor(tokens[0]);
    if (!color) return undefined;

    const positions: number[] = [];
    for (let i = 1; i < tokens.length; i++) {
      const tok = tokens[i];
      if (tok.endsWith('%')) {
        const pct = parseFloat(tok);
        if (!Number.isNaN(pct)) {
          positions.push(clamp(pct, 0, 100));
        }
      }
    }

    if (positions.length === 0) {
      expanded.push({ color });
    } else {
      for (const pos of positions) {
        expanded.push({ color, pos });
      }
    }
  }

  if (expanded.length < 2) return undefined;

  // Fill in CSS default stop positions (0% first, 100% last, linear interpolation between)
  if (expanded[0].pos === undefined) expanded[0].pos = 0;
  if (expanded[expanded.length - 1].pos === undefined) expanded[expanded.length - 1].pos = 100;

  let lastDefinedIdx = 0;
  for (let i = 1; i < expanded.length; i++) {
    if (expanded[i].pos !== undefined) {
      const gap = i - lastDefinedIdx;
      if (gap > 1) {
        const startPos = expanded[lastDefinedIdx].pos!;
        const endPos = expanded[i].pos!;
        for (let j = 1; j < gap; j++) {
          expanded[lastDefinedIdx + j].pos = Number(
            (startPos + ((endPos - startPos) * j) / gap).toFixed(2)
          );
        }
      }
      lastDefinedIdx = i;
    }
  }

  const stops: GradientStop[] = expanded.map((item) => {
    const stop: GradientStop = {
      color: item.color.hex,
      position: item.pos ?? 0,
    };
    if (item.color.opacity < 1) {
      stop.opacity = item.color.opacity;
    }
    return stop;
  });

  return { angleDeg, stops };
}

function parseLengthPx(token: string): number | undefined {
  const trimmed = token.trim().toLowerCase();
  if (trimmed === '0') return 0;
  const match = /^(-?\d*\.?\d+)(px|rem|em|pt)?$/.exec(trimmed);
  if (!match) return undefined;
  const val = parseFloat(match[1]);
  const unit = match[2] ?? 'px';
  switch (unit) {
    case 'px':
      return val;
    case 'rem':
    case 'em':
      return val * 16;
    case 'pt':
      return val * 2; // 0.5pt = 1px in 1920px coordinate space
    default:
      return undefined;
  }
}

/**
 * Parses CSS `border-radius` into pixels. Supports single values (`16px`),
 * multi-corner values (`28px 12px 24px 12px`), and percentage (`50%`) when
 * `wPx` and `hPx` are provided.
 */
export function parseBorderRadius(
  cssValue: string | undefined | null,
  wPx = 0,
  hPx = 0
): number | undefined {
  if (!cssValue) return undefined;
  const trimmed = cssValue.trim().toLowerCase();
  if (!trimmed || trimmed === 'none') return undefined;

  // Handle elliptical slash syntax "20px / 10px" by taking primary horizontal radii
  const primaryPart = trimmed.split('/')[0].trim();
  const tokens = primaryPart.split(/\s+/).filter(Boolean);
  if (tokens.length === 0) return undefined;

  const minDim = wPx > 0 && hPx > 0 ? Math.min(wPx, hPx) : 0;
  const values: number[] = [];

  for (const tok of tokens) {
    if (tok.endsWith('%')) {
      const pct = parseFloat(tok);
      if (Number.isNaN(pct)) continue;
      if (minDim > 0) {
        values.push((clamp(pct, 0, 100) / 100) * minDim);
      } else {
        return undefined;
      }
    } else {
      const px = parseLengthPx(tok);
      if (px !== undefined) {
        values.push(Math.max(0, px));
      }
    }
  }

  if (values.length === 0) return undefined;
  const avg = values.reduce((sum, v) => sum + v, 0) / values.length;
  return Number(avg.toFixed(2));
}

/**
 * Parses CSS `box-shadow` into an `ExtendedShadow`.
 * Supports both inline CSS (`0 8px 24px rgba(...)`) and computed style
 * order (`rgba(...) 0px 8px 24px 0px`). Ignores `inset` shadows.
 */
export function parseBoxShadow(
  cssValue: string | undefined | null,
  pxToPtFactor = DEFAULT_PX_TO_PT
): ExtendedShadow | undefined {
  if (!cssValue) return undefined;
  const trimmed = cssValue.trim();
  if (!trimmed || trimmed.toLowerCase() === 'none') return undefined;

  const layers = splitTopLevel(trimmed, ',');
  for (const layer of layers) {
    const tokens = tokenizeCssStop(layer);
    if (tokens.some((t) => t.toLowerCase() === 'inset')) {
      continue;
    }

    let color: ParsedColor | undefined;
    const lengths: number[] = [];

    for (const tok of tokens) {
      const parsedLength = parseLengthPx(tok);
      if (parsedLength !== undefined) {
        lengths.push(parsedLength);
        continue;
      }
      if (!color) {
        const parsedColor = parseCssColor(tok);
        if (parsedColor) {
          color = parsedColor;
        }
      }
    }

    if (lengths.length < 2) continue;

    const offsetXPx = lengths[0];
    const offsetYPx = lengths[1];
    const blurPx = lengths[2] ?? 0;
    const finalColor = color ?? { hex: '000000', opacity: 0.25 };

    return {
      color: finalColor.hex,
      opacity: finalColor.opacity,
      blurPt: Number((Math.max(0, blurPx) * pxToPtFactor).toFixed(2)),
      offsetXPx,
      offsetYPx,
    };
  }

  return undefined;
}

/**
 * Parses CSS `border` shorthand (e.g. `1.5px solid rgba(129, 140, 248, 0.32)`).
 */
export function parseBorder(
  cssValue: string | undefined | null,
  pxToPtFactor = DEFAULT_PX_TO_PT
): ExtendedShapeOptions['border'] | undefined {
  if (!cssValue) return undefined;
  const trimmed = cssValue.trim();
  if (!trimmed || trimmed.toLowerCase() === 'none') return undefined;

  const tokens = tokenizeCssStop(trimmed);
  let widthPx: number | undefined;
  let borderStyle = 'solid';
  let color: ParsedColor | undefined;

  for (const tok of tokens) {
    const lower = tok.toLowerCase();
    if (
      lower === 'none' ||
      lower === 'hidden' ||
      lower === 'solid' ||
      lower === 'dashed' ||
      lower === 'dotted' ||
      lower === 'double'
    ) {
      borderStyle = lower;
      continue;
    }

    const px = parseLengthPx(tok);
    if (px !== undefined) {
      widthPx = px;
      continue;
    }

    if (!color) {
      const parsed = parseCssColor(tok);
      if (parsed) color = parsed;
    }
  }

  if (borderStyle === 'none' || borderStyle === 'hidden') return undefined;
  if (!widthPx || widthPx <= 0) return undefined;

  const resolvedColor = color ?? { hex: '000000', opacity: 1 };
  if (resolvedColor.opacity <= 0) return undefined;

  const result: NonNullable<ExtendedShapeOptions['border']> = {
    color: resolvedColor.hex,
    widthPt: Number((widthPx * pxToPtFactor).toFixed(2)),
  };
  if (resolvedColor.opacity < 1) {
    result.opacity = resolvedColor.opacity;
  }
  return result;
}

/**
 * Parses CSS `filter: blur(Npx)` into `softEdgePt` for Aurora glow shapes.
 */
export function parseFilterBlur(
  cssValue: string | undefined | null,
  pxToPtFactor = DEFAULT_PX_TO_PT
): number | undefined {
  if (!cssValue) return undefined;
  const match = /blur\(\s*([^)]+)\s*\)/i.exec(cssValue);
  if (!match) return undefined;
  const px = parseLengthPx(match[1]);
  if (px === undefined || px <= 0) return undefined;
  return Number((px * pxToPtFactor).toFixed(2));
}

function getStyleProp(styles: CssStyleInput, camel: string, kebab: string): string | undefined {
  const record = styles as Record<string, unknown>;
  const val = record[camel] ?? record[kebab];
  if (typeof val === 'string' && val.trim() !== '') {
    return val.trim();
  }
  if (typeof (styles as CSSStyleDeclaration).getPropertyValue === 'function') {
    const propVal = (styles as CSSStyleDeclaration).getPropertyValue(kebab);
    if (propVal && propVal.trim() !== '') {
      return propVal.trim();
    }
  }
  return undefined;
}

/**
 * Converts CSS style properties (from either a plain object or `getComputedStyle()`)
 * and a 1920x1080 pixel bounding rect into `ExtendedShapeOptions`.
 */
export function parseCssToShapeOptions(
  styles: CssStyleInput,
  rect: RectPx,
  pxToPtFactor = DEFAULT_PX_TO_PT
): ExtendedShapeOptions {
  const background = getStyleProp(styles, 'background', 'background');
  const backgroundImage = getStyleProp(styles, 'backgroundImage', 'background-image');
  const backgroundColor = getStyleProp(styles, 'backgroundColor', 'background-color');
  const borderRadius = getStyleProp(styles, 'borderRadius', 'border-radius');
  const boxShadow = getStyleProp(styles, 'boxShadow', 'box-shadow');
  const borderStr = getStyleProp(styles, 'border', 'border');
  const filter = getStyleProp(styles, 'filter', 'filter');
  const opacityStr = getStyleProp(styles, 'opacity', 'opacity');

  const elementOpacity =
    opacityStr !== undefined && !Number.isNaN(parseFloat(opacityStr))
      ? clamp(parseFloat(opacityStr), 0, 1)
      : 1;

  const gradient =
    parseLinearGradient(backgroundImage) ?? parseLinearGradient(background);

  let fillColor: string | undefined;
  let fillOpacity: number | undefined;

  if (gradient) {
    if (elementOpacity < 1) {
      gradient.stops = gradient.stops.map((s) => ({
        ...s,
        opacity: roundOpacity((s.opacity ?? 1) * elementOpacity),
      }));
    }
  } else {
    const solidColor =
      parseCssColor(backgroundColor) ??
      (background && !background.includes('gradient(') ? parseCssColor(background) : undefined);

    if (solidColor) {
      fillColor = solidColor.hex;
      const combinedOpacity = roundOpacity(solidColor.opacity * elementOpacity);
      if (combinedOpacity < 1) {
        fillOpacity = combinedOpacity;
      }
    } else if (elementOpacity < 1) {
      fillOpacity = roundOpacity(elementOpacity);
    }
  }

  const cornerRadiusPx = parseBorderRadius(borderRadius, rect.wPx, rect.hPx);
  const shadow = parseBoxShadow(boxShadow, pxToPtFactor);
  const softEdgePt = parseFilterBlur(filter, pxToPtFactor);

  // Border fallback: if shorthand `border` is empty (common in getComputedStyle when sides differ),
  // check individual sides (top -> left -> right -> bottom).
  let border = parseBorder(borderStr, pxToPtFactor);
  if (!border) {
    for (const side of ['Top', 'Left', 'Right', 'Bottom'] as const) {
      const kebabSide = side.toLowerCase();
      const w = getStyleProp(styles, `border${side}Width`, `border-${kebabSide}-width`);
      const s = getStyleProp(styles, `border${side}Style`, `border-${kebabSide}-style`);
      const c = getStyleProp(styles, `border${side}Color`, `border-${kebabSide}-color`);
      if (w && s && c && s !== 'none') {
        border = parseBorder(`${w} ${s} ${c}`, pxToPtFactor);
        if (border) break;
      }
    }
  }

  if (border && elementOpacity < 1) {
    border.opacity = roundOpacity((border.opacity ?? 1) * elementOpacity);
  }

  let shapeType: PptxShapeName | undefined;
  if (borderRadius && borderRadius.trim() === '50%') {
    shapeType = 'ellipse';
  } else if (cornerRadiusPx !== undefined && cornerRadiusPx > 0) {
    shapeType = 'roundRect';
  } else if (cornerRadiusPx === 0) {
    shapeType = 'rect';
  }

  const options: ExtendedShapeOptions = {
    xPx: rect.xPx,
    yPx: rect.yPx,
    wPx: rect.wPx,
    hPx: rect.hPx,
  };

  if (shapeType) options.shapeType = shapeType;
  if (cornerRadiusPx !== undefined && shapeType !== 'ellipse') {
    options.cornerRadiusPx = cornerRadiusPx;
  }
  if (fillColor !== undefined) options.fillColor = fillColor;
  if (fillOpacity !== undefined) options.fillOpacity = fillOpacity;
  if (gradient) options.gradient = gradient;
  if (softEdgePt !== undefined) options.softEdgePt = softEdgePt;
  if (shadow) options.shadow = shadow;
  if (border) options.border = border;

  return options;
}

const GENERIC_FONT_FAMILIES = new Set([
  'sans-serif',
  'serif',
  'monospace',
  'ui-monospace',
  'ui-sans-serif',
  'ui-serif',
  'ui-rounded',
  'system-ui',
  '-apple-system',
  'blinkmacsystemfont',
  'cursive',
  'fantasy',
  'emoji',
  'math',
  'fangsong',
]);

export const DEFAULT_FONT_MAP: Record<string, string> = {
  'ui-monospace': 'Consolas',
  monospace: 'Consolas',
  'sfmono-regular': 'Consolas',
  'sf mono': 'Consolas',
  menlo: 'Consolas',
  monaco: 'Consolas',
  'liberation mono': 'Consolas',
  'courier new': 'Consolas',
};

/**
 * Extracts the primary font family from a CSS `font-family` stack, stripping
 * surrounding quotes, ignoring generic CSS fallbacks (`sans-serif`, `serif`,
 * `ui-monospace`, `monospace`, etc.), and applying `fontMap` (with sensible
 * defaults mapping monospace stacks to `Consolas`).
 */
export function resolveFontFamily(
  fontFamilyCss: string | undefined | null,
  fontMap?: Record<string, string>
): string {
  if (!fontFamilyCss || !fontFamilyCss.trim()) {
    return 'Arial';
  }

  const normalizedCustomMap = new Map<string, string>();
  if (fontMap) {
    for (const [key, val] of Object.entries(fontMap)) {
      normalizedCustomMap.set(key.trim().replace(/^["']|["']$/g, '').toLowerCase(), val);
    }
  }

  const tokens = splitTopLevel(fontFamilyCss, ',')
    .map((t) => t.trim().replace(/^["']|["']$/g, '').trim())
    .filter(Boolean);

  if (tokens.length === 0) return 'Arial';

  // 1. User-provided fontMap takes highest priority across tokens
  for (const token of tokens) {
    const mapped = normalizedCustomMap.get(token.toLowerCase());
    if (mapped) return mapped;
  }

  // 2. First non-generic font family (mapping known system monospace aliases to Consolas)
  for (const token of tokens) {
    const lower = token.toLowerCase();
    if (GENERIC_FONT_FAMILIES.has(lower)) continue;
    if (DEFAULT_FONT_MAP[lower]) return DEFAULT_FONT_MAP[lower];
    return token;
  }

  // 3. Stack contained only generic/system keywords -> check DEFAULT_FONT_MAP (e.g. ui-monospace -> Consolas)
  for (const token of tokens) {
    const lower = token.toLowerCase();
    if (DEFAULT_FONT_MAP[lower]) return DEFAULT_FONT_MAP[lower];
  }

  // 4. Generic serif fallback
  if (tokens.some((t) => t.toLowerCase() === 'serif' || t.toLowerCase() === 'ui-serif')) {
    return 'Georgia';
  }

  return 'Arial';
}

