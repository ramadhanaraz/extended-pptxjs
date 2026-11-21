import type PptxGenJS from 'pptxgenjs';
import {
  ExtendedPptxJS,
  type DonutChartOptions,
  type EmbeddedFont,
  type ExtendedGradient,
  type ExtendedShapeOptions,
  type GradientBarChartOptions,
  type GradientLineChartOptions,
  type GradientStop,
} from '../core/ExtendedPptxJS';
import {
  parseBoxShadow,
  parseCssColor,
  parseCssToShapeOptions,
  parseFilterBlur,
  parseLinearGradient,
  resolveFontFamily,
  type ParsedColor,
  type RectPx,
} from '../parsers/cssParser';
import {
  DEFAULT_RASTER_DPI_SCALE,
  hasBackdropBlur,
  isComplexGradientOrMesh,
  paintCssLayerToCanvas,
  parseRadialGradient,
  rasterizeSvgToPng,
  renderFrostedGlassElement,
  renderMeshBlurElement,
  renderSvgElement,
} from '../renderers/hybridAssetRenderer';

export type GlassMode = 'native-translucent' | 'raster-blur';

export interface HtmlSlideConverterOptions {
  dpiScale?: number;
  includeText?: boolean;
  baseWidthPx?: number;
  baseHeightPx?: number;
  glassMode?: GlassMode;
  fontMap?: Record<string, string>;
  embedFonts?: EmbeddedFont[];
}

interface SlideTransformMetrics {
  containerRect: DOMRect;
  cssScaleX: number;
  cssScaleY: number;
  designScaleX: number;
  designScaleY: number;
  designScale: number;
}

interface PaintItem {
  element: HTMLElement | SVGSVGElement;
  zIndex: number;
  isPositioned: boolean;
  domOrder: number;
}

interface ActiveBorderSide {
  side: 'top' | 'right' | 'bottom' | 'left';
  widthPx: number;
  style: string;
  color: ParsedColor;
}

const INLINE_PHRASING_TAGS = new Set([
  'SPAN',
  'STRONG',
  'B',
  'EM',
  'I',
  'U',
  'S',
  'MARK',
  'SMALL',
  'SUB',
  'SUP',
  'CODE',
  'A',
  'BR',
  'ABBR',
  'TIME',
]);

function getElementZOrder(el: Element): { zIndex: number; isPositioned: boolean } {
  const style = window.getComputedStyle(el);
  const pos = style.position;
  const isPositioned = pos !== 'static';
  const zRaw = style.zIndex;
  if (zRaw !== 'auto' && zRaw !== '') {
    const parsed = parseInt(zRaw, 10);
    if (!Number.isNaN(parsed)) {
      return { zIndex: parsed, isPositioned: true };
    }
  }
  return { zIndex: 0, isPositioned };
}

function sortChildrenInPaintOrder(parent: Element): Array<HTMLElement | SVGSVGElement> {
  const items: PaintItem[] = [];
  let order = 0;

  for (const child of Array.from(parent.children)) {
    if (child instanceof HTMLElement || child instanceof SVGSVGElement) {
      const { zIndex, isPositioned } = getElementZOrder(child);
      items.push({
        element: child,
        zIndex,
        isPositioned,
        domOrder: order++,
      });
    }
  }

  items.sort((a, b) => {
    const rank = (item: PaintItem): number => {
      if (item.zIndex < 0) return 0;
      if (!item.isPositioned && item.zIndex === 0) return 1;
      return 2;
    };
    const ra = rank(a);
    const rb = rank(b);
    if (ra !== rb) return ra - rb;
    if (a.zIndex !== b.zIndex) return a.zIndex - b.zIndex;
    return a.domOrder - b.domOrder;
  });

  return items.map((i) => i.element);
}

function isBackgroundClipText(style: CSSStyleDeclaration): boolean {
  const bgClip =
    style.backgroundClip ||
    style.getPropertyValue('-webkit-background-clip') ||
    style.getPropertyValue('background-clip') ||
    '';
  const textFill = style.getPropertyValue('-webkit-text-fill-color') || '';
  const hasTransparentTextFill =
    textFill === 'transparent' || textFill === 'rgba(0, 0, 0, 0)';
  return bgClip.includes('text') || hasTransparentTextFill;
}

function getActiveBorderSides(style: CSSStyleDeclaration): ActiveBorderSide[] {
  const sides: ActiveBorderSide[] = [];
  const entries: Array<['top' | 'right' | 'bottom' | 'left', 'Top' | 'Right' | 'Bottom' | 'Left']> = [
    ['top', 'Top'],
    ['right', 'Right'],
    ['bottom', 'Bottom'],
    ['left', 'Left'],
  ];

  for (const [side, cap] of entries) {
    const w = parseFloat(style[`border${cap}Width`]);
    const s = style[`border${cap}Style`];
    const c = parseCssColor(style[`border${cap}Color`]);
    if (!Number.isNaN(w) && w > 0 && s !== 'none' && s !== 'hidden' && c && c.opacity > 0) {
      sides.push({ side, widthPx: w, style: s, color: c });
    }
  }
  return sides;
}

function hasVisualBoxStyle(style: CSSStyleDeclaration): boolean {
  const bgColor = parseCssColor(style.backgroundColor);
  if (bgColor && bgColor.opacity > 0) return true;

  if (!isBackgroundClipText(style) && style.backgroundImage && style.backgroundImage !== 'none') {
    return true;
  }

  if (getActiveBorderSides(style).length > 0) return true;

  if (parseBoxShadow(style.boxShadow)) return true;

  return false;
}

function isBoldWeight(fontWeight: string): boolean {
  const num = parseInt(fontWeight, 10);
  if (!Number.isNaN(num)) return num >= 600;
  return fontWeight === 'bold' || fontWeight === 'bolder';
}

function mapTextAlign(textAlign: string): PptxGenJS.HAlign {
  switch (textAlign) {
    case 'center':
    case '-webkit-center':
      return 'center';
    case 'right':
    case 'end':
      return 'right';
    case 'justify':
      return 'justify';
    default:
      return 'left';
  }
}

function applyTextTransform(text: string, transform: string): string {
  switch (transform) {
    case 'uppercase':
      return text.toUpperCase();
    case 'lowercase':
      return text.toLowerCase();
    case 'capitalize':
      return text.replace(/\b\w/g, (c) => c.toUpperCase());
    default:
      return text;
  }
}

function extractTextLinearGradient(style: CSSStyleDeclaration): ExtendedGradient | undefined {
  if (!isBackgroundClipText(style)) return undefined;
  const bgImg = style.backgroundImage !== 'none' ? style.backgroundImage : style.background;
  return parseLinearGradient(bgImg);
}

function resolveTextColor(style: CSSStyleDeclaration): ParsedColor | undefined {
  if (isBackgroundClipText(style)) {
    const bgImg = style.backgroundImage !== 'none' ? style.backgroundImage : style.background;
    const lin = parseLinearGradient(bgImg);
    if (lin && lin.stops.length > 0) {
      return {
        hex: lin.stops[0].color,
        opacity: lin.stops[0].opacity ?? 1,
      };
    }
    const rad = bgImg ? parseRadialGradient(bgImg) : undefined;
    if (rad && rad.stops.length > 0) {
      return rad.stops[0].color;
    }
  }
  return parseCssColor(style.color);
}

function scaleShapeOptions(opts: ExtendedShapeOptions, designScale: number): ExtendedShapeOptions {
  if (Math.abs(designScale - 1) < 0.001) return opts;
  const scaled: ExtendedShapeOptions = { ...opts };
  if (scaled.cornerRadiusPx !== undefined) {
    scaled.cornerRadiusPx = Number((scaled.cornerRadiusPx * designScale).toFixed(2));
  }
  if (scaled.softEdgePt !== undefined) {
    scaled.softEdgePt = Number((scaled.softEdgePt * designScale).toFixed(2));
  }
  if (scaled.border) {
    scaled.border = {
      ...scaled.border,
      widthPt: Number((scaled.border.widthPt * designScale).toFixed(2)),
    };
  }
  if (scaled.shadow) {
    scaled.shadow = {
      ...scaled.shadow,
      blurPt: Number((scaled.shadow.blurPt * designScale).toFixed(2)),
      offsetXPx: Number((scaled.shadow.offsetXPx * designScale).toFixed(2)),
      offsetYPx: Number((scaled.shadow.offsetYPx * designScale).toFixed(2)),
    };
  }
  return scaled;
}

function parseSvgPointsAttr(pointsStr: string): Array<[number, number]> {
  const nums = pointsStr
    .trim()
    .split(/[\s,]+/)
    .map(Number)
    .filter((n) => !Number.isNaN(n));
  const pairs: Array<[number, number]> = [];
  for (let i = 0; i + 1 < nums.length; i += 2) {
    pairs.push([nums[i], nums[i + 1]]);
  }
  return pairs;
}

/**
 * Attempts to extract a native combo area+line chart definition from a simple
 * SVG area-line chart (`<polyline>` + `<polygon>` + optional `<linearGradient>`).
 */
export function extractAreaLineChartFromSvg(
  svgEl: SVGSVGElement,
  rect: RectPx
): GradientLineChartOptions | null {
  if (svgEl.querySelector('text')) return null;

  const polyline = svgEl.querySelector('polyline');
  const polygon = svgEl.querySelector('polygon');
  if (!polyline || !polygon) return null;

  const linePts = parseSvgPointsAttr(polyline.getAttribute('points') ?? '');
  const polyPts = parseSvgPointsAttr(polygon.getAttribute('points') ?? '');
  if (linePts.length < 3 || polyPts.length < 3) return null;

  const vbRaw = svgEl.getAttribute('viewBox');
  let vbHeight = rect.hPx;
  if (vbRaw) {
    const vb = vbRaw.trim().split(/[\s,]+/).map(Number);
    if (vb.length === 4 && vb[3] > 0) {
      vbHeight = vb[3];
    }
  }

  const maxPolyY = Math.max(...polyPts.map((p) => p[1]));
  const baselineY = maxPolyY > 0 ? maxPolyY : vbHeight;
  const scaleY = baselineY > 0 ? 100 / baselineY : 1;

  const values = linePts.map(([, y]) => Number(Math.max(0, (baselineY - y) * scaleY).toFixed(2)));
  const dataLabelsAttr = svgEl.getAttribute('data-labels');
  const parsedLabels = dataLabelsAttr
    ? dataLabelsAttr.split(',').map((s) => s.trim())
    : undefined;
  const labels =
    parsedLabels && parsedLabels.length === values.length
      ? parsedLabels
      : values.map((_, idx) => `P${idx + 1}`);

  const strokeAttr = polyline.getAttribute('stroke') ?? '#38BDF8';
  const parsedLineColor = parseCssColor(strokeAttr);
  const lineColor = parsedLineColor?.hex ?? '38BDF8';

  const strokeWidthAttr = parseFloat(polyline.getAttribute('stroke-width') ?? '4');
  const lineWidthPt = Number((Math.max(1, strokeWidthAttr) * 0.5).toFixed(2));

  let areaGradient: ExtendedGradient | undefined;
  const fillAttr = polygon.getAttribute('fill') ?? '';
  const urlMatch = /url\(\s*#([^)]+)\s*\)/.exec(fillAttr);
  if (urlMatch) {
    const gradEl = svgEl.querySelector(`#${CSS.escape(urlMatch[1])}`);
    if (gradEl && gradEl.tagName.toLowerCase() === 'lineargradient') {
      const stopEls = Array.from(gradEl.querySelectorAll('stop'));
      const stops: GradientStop[] = [];
      for (let i = 0; i < stopEls.length; i++) {
        const st = stopEls[i];
        const offsetRaw = st.getAttribute('offset') ?? `${(i / Math.max(1, stopEls.length - 1)) * 100}%`;
        const pos = offsetRaw.endsWith('%')
          ? parseFloat(offsetRaw)
          : parseFloat(offsetRaw) * 100;
        const stopColor = parseCssColor(st.getAttribute('stop-color') ?? `#${lineColor}`);
        const stopOpacityAttr = st.getAttribute('stop-opacity');
        const opacity =
          stopOpacityAttr !== null
            ? parseFloat(stopOpacityAttr)
            : stopColor?.opacity ?? 1;

        stops.push({
          color: stopColor?.hex ?? lineColor,
          position: Math.max(0, Math.min(100, pos)),
          opacity: Math.max(0, Math.min(1, opacity)),
        });
      }
      if (stops.length >= 2) {
        areaGradient = { angleDeg: 180, stops };
      }
    }
  }

  const circles = svgEl.querySelectorAll('circle');
  const showMarkers = circles.length > 0;
  let markerColor = lineColor;
  if (circles.length > 0) {
    const lastCircle = circles[circles.length - 1];
    const cFill = parseCssColor(lastCircle.getAttribute('fill') ?? '');
    if (cFill && cFill.opacity > 0) {
      markerColor = cFill.hex;
    }
  }

  return {
    xPx: rect.xPx,
    yPx: rect.yPx,
    wPx: rect.wPx,
    hPx: rect.hPx,
    seriesName: svgEl.getAttribute('data-series-name') ?? 'Series 1',
    labels,
    values,
    lineColor,
    lineWidthPt,
    showMarkers,
    markerSize: 7,
    markerColor,
    areaGradient,
    valAxisMin: 0,
    valAxisMax: 100,
  };
}

export class HtmlSlideConverter {
  private readonly dpiScale: number;
  private readonly includeText: boolean;
  private readonly baseWidthPx: number;
  private readonly baseHeightPx: number;
  private readonly glassMode: GlassMode;
  private readonly fontMap?: Record<string, string>;

  constructor(
    private readonly ext: ExtendedPptxJS,
    options: HtmlSlideConverterOptions = {}
  ) {
    this.dpiScale = options.dpiScale ?? DEFAULT_RASTER_DPI_SCALE;
    this.includeText = options.includeText ?? true;
    this.baseWidthPx = options.baseWidthPx ?? ext.baseWidthPx;
    this.baseHeightPx = options.baseHeightPx ?? ext.baseHeightPx;
    this.glassMode = options.glassMode ?? 'native-translucent';
    this.fontMap = options.fontMap;

    if (options.embedFonts?.length) {
      this.ext.embedFonts(options.embedFonts);
    }
  }

  private computeMetrics(slideElement: HTMLElement): SlideTransformMetrics {
    const containerRect = slideElement.getBoundingClientRect();
    const offsetW = slideElement.offsetWidth || this.baseWidthPx;
    const offsetH = slideElement.offsetHeight || this.baseHeightPx;
    const cssScaleX = containerRect.width > 0 ? containerRect.width / offsetW : 1;
    const cssScaleY = containerRect.height > 0 ? containerRect.height / offsetH : 1;
    const designScaleX = this.baseWidthPx / offsetW;
    const designScaleY = this.baseHeightPx / offsetH;

    return {
      containerRect,
      cssScaleX,
      cssScaleY,
      designScaleX,
      designScaleY,
      designScale: Math.min(designScaleX, designScaleY),
    };
  }

  private toSlideRect(elRect: DOMRect, metrics: SlideTransformMetrics): RectPx {
    const unscaledX = (elRect.left - metrics.containerRect.left) / metrics.cssScaleX;
    const unscaledY = (elRect.top - metrics.containerRect.top) / metrics.cssScaleY;
    const unscaledW = elRect.width / metrics.cssScaleX;
    const unscaledH = elRect.height / metrics.cssScaleY;

    return {
      xPx: Number((unscaledX * metrics.designScaleX).toFixed(2)),
      yPx: Number((unscaledY * metrics.designScaleY).toFixed(2)),
      wPx: Number((unscaledW * metrics.designScaleX).toFixed(2)),
      hPx: Number((unscaledH * metrics.designScaleY).toFixed(2)),
    };
  }

  private createBackdropCanvas(): HTMLCanvasElement | undefined {
    if (typeof document === 'undefined') return undefined;
    const canvas = document.createElement('canvas');
    canvas.width = this.baseWidthPx;
    canvas.height = this.baseHeightPx;
    return canvas;
  }

  private async withVisibleSlide<T>(
    slideElement: HTMLElement,
    fn: () => Promise<T>
  ): Promise<T> {
    const restores: Array<() => void> = [];
    const chain: HTMLElement[] = [];
    let cur: HTMLElement | null = slideElement;
    while (cur && cur !== document.body && cur !== document.documentElement) {
      chain.push(cur);
      cur = cur.parentElement;
    }

    // Unhide from outermost hidden ancestor down to slideElement
    for (let i = chain.length - 1; i >= 0; i--) {
      const node = chain[i];
      const cs = window.getComputedStyle(node);
      if (cs.display === 'none') {
        const prevVal = node.style.getPropertyValue('display');
        const prevPrio = node.style.getPropertyPriority('display');
        const inferred =
          cs.gridTemplateColumns && cs.gridTemplateColumns !== 'none'
            ? 'grid'
            : cs.flexDirection !== 'row' ||
                cs.justifyContent !== 'normal' ||
                cs.alignItems !== 'normal'
              ? 'flex'
              : 'block';
        node.style.setProperty('display', inferred, 'important');
        restores.push(() => {
          if (prevVal) {
            node.style.setProperty('display', prevVal, prevPrio);
          } else {
            node.style.removeProperty('display');
          }
        });
      }
    }

    if (restores.length > 0) {
      // Force synchronous layout flush after revealing hidden slide
      void slideElement.offsetWidth;
    }

    try {
      return await fn();
    } finally {
      for (let i = restores.length - 1; i >= 0; i--) {
        restores[i]();
      }
    }
  }

  private emitDeclarativeChartAttr(
    el: Element,
    slide: PptxGenJS.Slide,
    rect: RectPx
  ): boolean {
    const raw = el.getAttribute('data-pptx-chart');
    if (!raw) return false;

    try {
      const parsed = JSON.parse(raw) as Partial<
        GradientLineChartOptions & GradientBarChartOptions & DonutChartOptions
      > & {
        type?: 'gradient-line' | 'line' | 'bar' | 'col' | 'donut' | 'doughnut';
        color?: string;
        gradientStops?: GradientStop[];
      };

      if (!Array.isArray(parsed.labels) || !Array.isArray(parsed.values)) {
        return false;
      }

      const chartType = parsed.type ?? 'gradient-line';
      const labels = parsed.labels.map(String);
      const values = parsed.values.map(Number);

      if (chartType === 'donut' || chartType === 'doughnut') {
        this.ext.addDonutChart(slide, {
          xPx: rect.xPx,
          yPx: rect.yPx,
          wPx: rect.wPx,
          hPx: rect.hPx,
          seriesName: parsed.seriesName,
          labels,
          values,
          colors: parsed.colors,
          holeSize: parsed.holeSize,
          showLegend: parsed.showLegend,
          showDataLabels: parsed.showDataLabels,
        });
        return true;
      }

      const rawColor = parsed.barColor ?? parsed.lineColor ?? parsed.color ?? '#38BDF8';
      const resolvedHex = parseCssColor(rawColor)?.hex ?? rawColor.replace('#', '');

      let customGradient = parsed.gradient ?? parsed.areaGradient;
      if (!customGradient && Array.isArray(parsed.gradientStops)) {
        customGradient = {
          angleDeg: 180,
          stops: parsed.gradientStops.map((s) => ({
            ...s,
            color: (parseCssColor(s.color)?.hex ?? s.color).replace('#', ''),
          })),
        };
      }

      if (chartType === 'bar' || chartType === 'col') {
        this.ext.addGradientBarChart(slide, {
          xPx: rect.xPx,
          yPx: rect.yPx,
          wPx: rect.wPx,
          hPx: rect.hPx,
          seriesName: parsed.seriesName,
          labels,
          values,
          barDir: chartType === 'bar' ? 'bar' : 'col',
          barColor: resolvedHex,
          gradient: customGradient,
          showDataLabels: parsed.showDataLabels,
          dataLabelColor: parsed.dataLabelColor,
          dataLabelPosition: parsed.dataLabelPosition,
          showGridLines: parsed.showGridLines,
          showAxes: parsed.showAxes,
          showLegend: parsed.showLegend,
          barGapWidthPct: parsed.barGapWidthPct,
          valAxisMin: parsed.valAxisMin,
          valAxisMax: parsed.valAxisMax,
        });
        return true;
      }

      this.ext.addGradientLineChart(slide, {
        xPx: rect.xPx,
        yPx: rect.yPx,
        wPx: rect.wPx,
        hPx: rect.hPx,
        seriesName: parsed.seriesName,
        labels,
        values,
        lineColor: resolvedHex,
        lineWidthPt: parsed.lineWidthPt,
        showMarkers: parsed.showMarkers,
        markerSize: parsed.markerSize,
        markerColor: parsed.markerColor
          ? (parseCssColor(parsed.markerColor)?.hex ?? parsed.markerColor.replace('#', ''))
          : undefined,
        areaGradient: customGradient,
        smooth: parsed.smooth,
        showLegend: parsed.showLegend,
        valAxisMin: parsed.valAxisMin,
        valAxisMax: parsed.valAxisMax,
      });
      return true;
    } catch {
      return false;
    }
  }

  private emitSingleSidedBorderLines(
    slide: PptxGenJS.Slide,
    rect: RectPx,
    sides: ActiveBorderSide[],
    elementOpacity: number,
    designScale: number
  ): void {
    for (const s of sides) {
      const widthPt = Number((s.widthPx * designScale * 0.5).toFixed(2));
      const combinedOpacity = s.color.opacity * elementOpacity;
      const transparency =
        combinedOpacity < 1 ? Math.round((1 - combinedOpacity) * 100) : 0;
      const dashType: PptxGenJS.ShapeLineProps['dashType'] =
        s.style === 'dashed' ? 'dash' : s.style === 'dotted' ? 'sysDot' : 'solid';

      let xPx = rect.xPx;
      let yPx = rect.yPx;
      let wPx = rect.wPx;
      let hPx = rect.hPx;

      switch (s.side) {
        case 'top':
          hPx = 0;
          break;
        case 'bottom':
          yPx = rect.yPx + rect.hPx;
          hPx = 0;
          break;
        case 'left':
          wPx = 0;
          break;
        case 'right':
          xPx = rect.xPx + rect.wPx;
          wPx = 0;
          break;
      }

      slide.addShape(this.ext.pptx.ShapeType.line, {
        x: this.ext.pxToIn(xPx),
        y: this.ext.pxToIn(yPx),
        w: this.ext.pxToIn(wPx),
        h: this.ext.pxToIn(hPx),
        line: {
          color: s.color.hex,
          width: widthPt,
          transparency,
          dashType,
        },
      });
    }
  }

  private emitBoxWithSelectiveBorders(
    slide: PptxGenJS.Slide,
    style: CSSStyleDeclaration,
    rect: RectPx,
    metrics: SlideTransformMetrics,
    forceTranslucentGlass = false
  ): void {
    const activeSides = getActiveBorderSides(style);
    const elOpacityRaw = parseFloat(style.opacity);
    const elementOpacity = !Number.isNaN(elOpacityRaw)
      ? Math.max(0, Math.min(1, elOpacityRaw))
      : 1;

    const rawShapeOpts = parseCssToShapeOptions(style, rect);
    if (isBackgroundClipText(style)) {
      delete rawShapeOpts.gradient;
    }

    const blurPt = parseFilterBlur(style.filter);
    if (blurPt !== undefined && blurPt > 0) {
      rawShapeOpts.softEdgePt = blurPt;
    }

    if (forceTranslucentGlass) {
      if (!rawShapeOpts.shapeType) {
        rawShapeOpts.shapeType = 'roundRect';
      }
      if (rawShapeOpts.fillOpacity === undefined && !rawShapeOpts.gradient) {
        rawShapeOpts.fillOpacity = 0.72;
      }
    }

    // Only keep a uniform 4-sided shape border when ALL 4 sides match;
    // for 1-3 sides emit lines only, and for asymmetric 4-sided HUD/TOC borders
    // keep the thinnest base border on the shape and emit accent edges as native lines.
    let accentLineSides: ActiveBorderSide[] = [];
    if (activeSides.length > 0 && activeSides.length < 4) {
      delete rawShapeOpts.border;
      accentLineSides = activeSides;
    } else if (activeSides.length === 4) {
      const ref = activeSides[0];
      const isUniform4 = activeSides.every(
        (s) =>
          Math.abs(s.widthPx - ref.widthPx) <= 0.25 &&
          s.color.hex === ref.color.hex &&
          Math.abs(s.color.opacity - ref.color.opacity) <= 0.02 &&
          s.style === ref.style
      );
      if (!isUniform4) {
        const minSide = activeSides.reduce((min, s) =>
          s.widthPx < min.widthPx ? s : min
        );
        rawShapeOpts.border = {
          color: minSide.color.hex,
          widthPt: Number((minSide.widthPx * 0.5).toFixed(2)),
          opacity: minSide.color.opacity * elementOpacity < 1
            ? Number((minSide.color.opacity * elementOpacity).toFixed(4))
            : undefined,
        };
        accentLineSides = activeSides.filter(
          (s) =>
            Math.abs(s.widthPx - minSide.widthPx) > 0.25 ||
            s.color.hex !== minSide.color.hex
        );
      }
    } else if (activeSides.length === 0 && forceTranslucentGlass) {
      rawShapeOpts.border = {
        color: 'FFFFFF',
        widthPt: 1,
        opacity: 0.22,
      };
    }

    const shapeOpts = scaleShapeOptions(rawShapeOpts, metrics.designScale);

    const hasFillOrShadow =
      Boolean(shapeOpts.gradient) ||
      (shapeOpts.fillColor !== undefined && (shapeOpts.fillOpacity ?? 1) > 0) ||
      Boolean(shapeOpts.shadow) ||
      Boolean(shapeOpts.border);

    if (hasFillOrShadow) {
      this.ext.addExtendedShape(slide, shapeOpts);
    }

    if (accentLineSides.length > 0) {
      this.emitSingleSidedBorderLines(
        slide,
        rect,
        accentLineSides,
        elementOpacity,
        metrics.designScale
      );
    }
  }

  private async emitPseudoElement(
    parentEl: HTMLElement,
    pseudoName: '::before' | '::after',
    slide: PptxGenJS.Slide,
    parentRect: RectPx,
    metrics: SlideTransformMetrics,
    backdropCanvas: HTMLCanvasElement | undefined
  ): Promise<void> {
    const pseudo = window.getComputedStyle(parentEl, pseudoName);
    const content = pseudo.content;
    if (!content || content === 'none' || pseudo.display === 'none' || pseudo.visibility === 'hidden') {
      return;
    }
    const opacity = parseFloat(pseudo.opacity);
    if (!Number.isNaN(opacity) && opacity <= 0) return;

    const parseOptPx = (val: string): number | undefined => {
      if (!val || val === 'auto') return undefined;
      const n = parseFloat(val);
      return Number.isNaN(n) ? undefined : n;
    };

    const pWidth = parseOptPx(pseudo.width);
    const pHeight = parseOptPx(pseudo.height);
    const pLeft = parseOptPx(pseudo.left);
    const pRight = parseOptPx(pseudo.right);
    const pTop = parseOptPx(pseudo.top);
    const pBottom = parseOptPx(pseudo.bottom);

    let wPx: number | undefined;
    let hPx: number | undefined;
    let xPx = parentRect.xPx;
    let yPx = parentRect.yPx;

    if (pseudo.position === 'absolute') {
      if (pWidth !== undefined) {
        wPx = pWidth * metrics.designScaleX;
      } else if (pLeft !== undefined && pRight !== undefined) {
        wPx = parentRect.wPx - (pLeft + pRight) * metrics.designScaleX;
      }

      if (pHeight !== undefined) {
        hPx = pHeight * metrics.designScaleY;
      } else if (pTop !== undefined && pBottom !== undefined) {
        hPx = parentRect.hPx - (pTop + pBottom) * metrics.designScaleY;
      }

      if (pLeft !== undefined) {
        xPx = parentRect.xPx + pLeft * metrics.designScaleX;
      } else if (pRight !== undefined && wPx !== undefined) {
        xPx = parentRect.xPx + parentRect.wPx - pRight * metrics.designScaleX - wPx;
      }

      if (pTop !== undefined) {
        yPx = parentRect.yPx + pTop * metrics.designScaleY;
      } else if (pBottom !== undefined && hPx !== undefined) {
        yPx = parentRect.yPx + parentRect.hPx - pBottom * metrics.designScaleY - hPx;
      }
    } else if (pWidth !== undefined && pHeight !== undefined) {
      wPx = pWidth * metrics.designScaleX;
      hPx = pHeight * metrics.designScaleY;
    }

    if (!wPx || !hPx || wPx <= 0.5 || hPx <= 0.5) return;
    const pseudoRect: RectPx = {
      xPx: Number(xPx.toFixed(2)),
      yPx: Number(yPx.toFixed(2)),
      wPx: Number(wPx.toFixed(2)),
      hPx: Number(hPx.toFixed(2)),
    };

    if (isComplexGradientOrMesh(pseudo)) {
      const asset = await renderMeshBlurElement(pseudo, pseudoRect, this.dpiScale);
      slide.addImage({
        data: asset.dataUrl,
        x: this.ext.pxToIn(asset.xPx),
        y: this.ext.pxToIn(asset.yPx),
        w: this.ext.pxToIn(asset.wPx),
        h: this.ext.pxToIn(asset.hPx),
      });
      const ctx = backdropCanvas?.getContext('2d');
      if (ctx) paintCssLayerToCanvas(ctx, pseudo, pseudoRect);
    } else if (hasVisualBoxStyle(pseudo)) {
      this.emitBoxWithSelectiveBorders(slide, pseudo, pseudoRect, metrics, false);
      const ctx = backdropCanvas?.getContext('2d');
      if (ctx) paintCssLayerToCanvas(ctx, pseudo, pseudoRect);
    }

    const strippedText = content.replace(/^["']|["']$/g, '').trim();
    if (strippedText && !/^[\uE000-\uF8FF\s]+$/.test(strippedText)) {
      const fontColor = resolveTextColor(pseudo);
      const fontSizePx = (parseFloat(pseudo.fontSize) || 16) * metrics.designScale;
      slide.addText(applyTextTransform(strippedText, pseudo.textTransform), {
        x: this.ext.pxToIn(pseudoRect.xPx),
        y: this.ext.pxToIn(pseudoRect.yPx),
        w: this.ext.pxToIn(pseudoRect.wPx * 1.14),
        h: this.ext.pxToIn(Math.max(pseudoRect.hPx, fontSizePx * 1.2)),
        fontFace: resolveFontFamily(pseudo.fontFamily, this.fontMap),
        fontSize: Number((fontSizePx * 0.5).toFixed(2)),
        bold: isBoldWeight(pseudo.fontWeight),
        italic: pseudo.fontStyle === 'italic',
        color: fontColor?.hex ?? '000000',
        align: mapTextAlign(pseudo.textAlign),
        valign: 'middle',
        margin: 0,
        wrap: false,
        fit: 'shrink',
      });
    }
  }

  private processTableElement(
    tableEl: HTMLTableElement,
    slide: PptxGenJS.Slide,
    metrics: SlideTransformMetrics,
    tableRect: RectPx
  ): void {
    const rows: PptxGenJS.TableRow[] = [];
    const rowHeightsIn: number[] = [];
    let colWidthsIn: number[] | undefined;

    for (const tr of Array.from(tableEl.rows)) {
      const trRect = this.toSlideRect(tr.getBoundingClientRect(), metrics);
      if (trRect.hPx > 0) {
        rowHeightsIn.push(this.ext.pxToIn(trRect.hPx));
      }

      const trStyle = window.getComputedStyle(tr);
      const rowCells: PptxGenJS.TableCell[] = [];
      const rowColWidths: number[] = [];

      for (const cell of Array.from(tr.cells)) {
        const cellStyle = window.getComputedStyle(cell);
        const cellRect = this.toSlideRect(cell.getBoundingClientRect(), metrics);
        if (cell.colSpan === 1 && cellRect.wPx > 0) {
          rowColWidths.push(this.ext.pxToIn(cellRect.wPx));
        }

        const rawText = (cell.innerText || cell.textContent || '').replace(/\s+/g, ' ').trim();
        const text = applyTextTransform(rawText, cellStyle.textTransform);

        const fontColor = resolveTextColor(cellStyle);
        const fontSizePx = (parseFloat(cellStyle.fontSize) || 16) * metrics.designScale;
        const fontSizePt = Number((fontSizePx * 0.5).toFixed(2));
        const fontFace = resolveFontFamily(cellStyle.fontFamily, this.fontMap);
        const bold = cell.tagName === 'TH' || isBoldWeight(cellStyle.fontWeight);
        const italic = cellStyle.fontStyle === 'italic';
        const align = mapTextAlign(cellStyle.textAlign);

        let bgColor = parseCssColor(cellStyle.backgroundColor);
        if (!bgColor || bgColor.opacity <= 0) {
          bgColor = parseCssColor(trStyle.backgroundColor);
        }

        const fill: PptxGenJS.ShapeFillProps | undefined =
          bgColor && bgColor.opacity > 0
            ? {
                color: bgColor.hex,
                transparency: Math.round((1 - bgColor.opacity) * 100),
              }
            : undefined;

        const mapBorder = (
          side: 'Top' | 'Right' | 'Bottom' | 'Left'
        ): PptxGenJS.BorderProps => {
          const wPx = parseFloat(cellStyle[`border${side}Width`]);
          const s = cellStyle[`border${side}Style`];
          const c = parseCssColor(cellStyle[`border${side}Color`]);
          if (!Number.isNaN(wPx) && wPx > 0 && s !== 'none' && s !== 'hidden' && c && c.opacity > 0) {
            return {
              type: 'solid',
              color: c.hex,
              pt: Number((wPx * metrics.designScale * 0.5).toFixed(2)),
            };
          }
          return { type: 'none' };
        };

        const padT = Number(((parseFloat(cellStyle.paddingTop) || 0) * metrics.designScale * 0.5).toFixed(1));
        const padR = Number(((parseFloat(cellStyle.paddingRight) || 0) * metrics.designScale * 0.5).toFixed(1));
        const padB = Number(((parseFloat(cellStyle.paddingBottom) || 0) * metrics.designScale * 0.5).toFixed(1));
        const padL = Number(((parseFloat(cellStyle.paddingLeft) || 0) * metrics.designScale * 0.5).toFixed(1));

        const cellOpts: PptxGenJS.TableCellProps = {
          fontFace,
          fontSize: fontSizePt,
          bold,
          italic,
          color: fontColor?.hex ?? '000000',
          align,
          valign: 'middle',
          margin: [padT, padR, padB, padL],
          border: [mapBorder('Top'), mapBorder('Right'), mapBorder('Bottom'), mapBorder('Left')],
        };

        if (fill) cellOpts.fill = fill;
        if (cell.colSpan > 1) cellOpts.colspan = cell.colSpan;
        if (cell.rowSpan > 1) cellOpts.rowspan = cell.rowSpan;

        rowCells.push({
          text,
          options: cellOpts,
        });
      }

      if (!colWidthsIn && rowColWidths.length > 0 && rowCells.every((c) => !c.options?.colspan)) {
        colWidthsIn = rowColWidths;
      }

      if (rowCells.length > 0) {
        rows.push(rowCells);
      }
    }

    if (rows.length === 0) return;

    const tableProps: PptxGenJS.TableProps = {
      x: this.ext.pxToIn(tableRect.xPx),
      y: this.ext.pxToIn(tableRect.yPx),
      w: this.ext.pxToIn(tableRect.wPx),
    };
    if (colWidthsIn) tableProps.colW = colWidthsIn;
    if (rowHeightsIn.length === rows.length) tableProps.rowH = rowHeightsIn;

    slide.addTable(rows, tableProps);
  }

  private isSingleLineRange(
    range: Range,
    style: CSSStyleDeclaration,
    metrics: SlideTransformMetrics,
    hasExplicitLineBreak: boolean
  ): boolean {
    if (hasExplicitLineBreak) return false;

    const fontSizeUnscaled = parseFloat(style.fontSize) || 16;
    const lineHeightRaw = parseFloat(style.lineHeight);
    const lineHeightUnscaled = !Number.isNaN(lineHeightRaw)
      ? lineHeightRaw
      : fontSizeUnscaled * 1.25;

    const bounding = range.getBoundingClientRect();
    const unscaledHeight = bounding.height / metrics.cssScaleY;
    if (unscaledHeight > lineHeightUnscaled * 1.55) {
      return false;
    }

    const rects = Array.from(range.getClientRects()).filter(
      (r) => r.width > 0.5 && r.height > 0.5
    );
    if (rects.length <= 1) return true;

    const lineTops: number[] = [];
    const tolerance = Math.max(4, fontSizeUnscaled * metrics.cssScaleY * 0.45);
    for (const r of rects) {
      if (!lineTops.some((t) => Math.abs(t - r.top) <= tolerance)) {
        lineTops.push(r.top);
      }
      if (lineTops.length > 1) return false;
    }
    return true;
  }

  private computeTextBoxRect(
    rawRect: RectPx,
    align: PptxGenJS.HAlign,
    isSingleLine: boolean,
    fontSizePx: number
  ): { xPx: number; yPx: number; wPx: number; hPx: number } {
    let xPx = rawRect.xPx;
    let wPx = rawRect.wPx;

    if (isSingleLine) {
      const expandedW = Math.max(rawRect.wPx * 1.14, rawRect.wPx + 4);
      const extra = expandedW - rawRect.wPx;
      if (align === 'center') {
        xPx = Math.max(0, rawRect.xPx - extra / 2);
      } else if (align === 'right') {
        xPx = Math.max(0, rawRect.xPx - extra);
      }
      wPx = Math.min(expandedW, Math.max(1, this.baseWidthPx - xPx));
    } else {
      wPx = Math.min(rawRect.wPx + 6, Math.max(1, this.baseWidthPx - xPx));
    }

    const hPx = Math.max(rawRect.hPx, fontSizePx * 1.2);
    return { xPx, yPx: rawRect.yPx, wPx, hPx };
  }

  private isRichTextContainer(el: HTMLElement): boolean {
    if (el.children.length === 0) return false;

    const containerStyle = window.getComputedStyle(el);
    if (
      containerStyle.display === 'flex' ||
      containerStyle.display === 'inline-flex' ||
      containerStyle.display === 'grid' ||
      containerStyle.display === 'inline-grid'
    ) {
      return false;
    }

    const checkNode = (parent: HTMLElement): boolean => {
      for (const child of Array.from(parent.children)) {
        if (!(child instanceof HTMLElement)) return false;
        if (!INLINE_PHRASING_TAGS.has(child.tagName)) return false;
        if (child.hasAttribute('data-pptx-chart') || child.hasAttribute('data-pptx-raster')) {
          return false;
        }

        const cs = window.getComputedStyle(child);
        if (cs.display === 'none' || cs.visibility === 'hidden') continue;
        if (child.tagName !== 'BR' && hasVisualBoxStyle(cs)) {
          return false;
        }
        if (
          child.tagName !== 'BR' &&
          (cs.display === 'block' ||
            cs.display === 'flex' ||
            cs.display === 'inline-flex' ||
            cs.display === 'grid' ||
            cs.display === 'inline-grid' ||
            cs.display === 'list-item' ||
            cs.display === 'table' ||
            cs.display === 'table-cell' ||
            cs.position === 'absolute' ||
            cs.position === 'fixed')
        ) {
          return false;
        }
        if (
          child.tagName !== 'BR' &&
          (cs.display === 'inline-block') &&
          (parseFloat(cs.paddingLeft) > 0 || parseFloat(cs.paddingRight) > 0)
        ) {
          return false;
        }
        if (!checkNode(child)) return false;
      }
      return true;
    };

    return checkNode(el);
  }

  private collectInlineRuns(
    node: Node,
    parentEl: HTMLElement,
    slide: PptxGenJS.Slide,
    metrics: SlideTransformMetrics,
    runs: PptxGenJS.TextProps[],
    state: { hasLineBreak: boolean }
  ): void {
    if (node.nodeType === Node.TEXT_NODE) {
      const raw = node.textContent ?? '';
      const collapsed = raw.replace(/\s+/g, ' ');
      if (!collapsed || collapsed === ' ') {
        if (
          collapsed === ' ' &&
          runs.length > 0 &&
          !runs[runs.length - 1].options?.breakLine &&
          !runs[runs.length - 1].text?.endsWith(' ')
        ) {
          runs[runs.length - 1].text = `${runs[runs.length - 1].text} `;
        }
        return;
      }

      const style = window.getComputedStyle(parentEl);
      if (style.display === 'none' || style.visibility === 'hidden') return;
      const elOpacity = parseFloat(style.opacity);
      if (!Number.isNaN(elOpacity) && elOpacity <= 0) return;

      const text = applyTextTransform(collapsed, style.textTransform);
      const textGrad = extractTextLinearGradient(style);
      const fontColor = textGrad ? undefined : resolveTextColor(style);
      const runColor = textGrad
        ? this.ext.registerTextGradient(slide, textGrad)
        : (fontColor?.hex ?? '000000');

      const fontSizePx = (parseFloat(style.fontSize) || 16) * metrics.designScale;
      const fontSizePt = Number((fontSizePx * 0.5).toFixed(2));
      const fontFace = resolveFontFamily(style.fontFamily, this.fontMap);
      const bold = isBoldWeight(style.fontWeight);
      const italic = style.fontStyle === 'italic';
      const underline = style.textDecorationLine.includes('underline')
        ? ({ style: 'sng' } as const)
        : undefined;
      const letterSpacingPx = parseFloat(style.letterSpacing);
      const charSpacing =
        !Number.isNaN(letterSpacingPx) && letterSpacingPx !== 0
          ? Number((letterSpacingPx * metrics.designScale * 0.5).toFixed(2))
          : undefined;
      const combinedOpacity =
        textGrad
          ? 1
          : (fontColor?.opacity ?? 1) * (Number.isNaN(elOpacity) ? 1 : elOpacity);
      const transparency =
        combinedOpacity < 1 ? Math.round((1 - combinedOpacity) * 100) : 0;

      let finalText = text;
      if (runs.length === 0 || runs[runs.length - 1].options?.breakLine) {
        finalText = finalText.replace(/^\s+/, '');
      }
      if (!finalText) return;

      runs.push({
        text: finalText,
        options: {
          fontFace,
          fontSize: fontSizePt,
          bold,
          italic,
          underline,
          color: runColor,
          transparency,
          charSpacing,
        },
      });
      return;
    }

    if (node instanceof HTMLElement) {
      if (node.tagName === 'BR') {
        state.hasLineBreak = true;
        if (runs.length > 0) {
          runs[runs.length - 1].options = {
            ...runs[runs.length - 1].options,
            breakLine: true,
          };
        } else {
          runs.push({ text: '', options: { breakLine: true } });
        }
        return;
      }

      const cs = window.getComputedStyle(node);
      if (cs.display === 'none' || cs.visibility === 'hidden') return;

      for (const child of Array.from(node.childNodes)) {
        this.collectInlineRuns(child, node, slide, metrics, runs, state);
      }
    }
  }

  private emitRichTextBlock(
    el: HTMLElement,
    slide: PptxGenJS.Slide,
    metrics: SlideTransformMetrics
  ): boolean {
    if (!this.includeText) return false;
    if (!this.isRichTextContainer(el)) return false;

    const style = window.getComputedStyle(el);
    if (style.display === 'none' || style.visibility === 'hidden') return true;
    const elOpacity = parseFloat(style.opacity);
    if (!Number.isNaN(elOpacity) && elOpacity <= 0) return true;

    const runs: PptxGenJS.TextProps[] = [];
    const state = { hasLineBreak: false };
    for (const child of Array.from(el.childNodes)) {
      this.collectInlineRuns(child, el, slide, metrics, runs, state);
    }

    if (runs.length === 0) return true;
    const lastRun = runs[runs.length - 1];
    if (lastRun.text) {
      lastRun.text = lastRun.text.replace(/\s+$/, '');
    }
    if (runs.every((r) => !r.text)) return true;

    const range = document.createRange();
    range.selectNodeContents(el);
    const bounding = range.getBoundingClientRect();
    const rawRect = this.toSlideRect(bounding, metrics);
    if (rawRect.wPx <= 0.5 || rawRect.hPx <= 0.5) return true;

    const align = mapTextAlign(style.textAlign);
    const fontSizePx = (parseFloat(style.fontSize) || 16) * metrics.designScale;
    const isSingleLine = this.isSingleLineRange(range, style, metrics, state.hasLineBreak);
    const box = this.computeTextBoxRect(rawRect, align, isSingleLine, fontSizePx);

    slide.addText(runs, {
      x: this.ext.pxToIn(box.xPx),
      y: this.ext.pxToIn(box.yPx),
      w: this.ext.pxToIn(box.wPx),
      h: this.ext.pxToIn(box.hPx),
      align,
      valign: isSingleLine ? 'middle' : 'top',
      margin: 0,
      wrap: !isSingleLine,
      fit: isSingleLine ? 'shrink' : 'none',
    });

    return true;
  }

  private emitDirectTextNodes(
    el: HTMLElement,
    slide: PptxGenJS.Slide,
    metrics: SlideTransformMetrics
  ): void {
    if (!this.includeText) return;

    const style = window.getComputedStyle(el);
    if (style.display === 'none' || style.visibility === 'hidden') return;

    const elOpacity = parseFloat(style.opacity);
    if (!Number.isNaN(elOpacity) && elOpacity <= 0) return;

    const textGrad = extractTextLinearGradient(style);
    const fontColor = textGrad ? undefined : resolveTextColor(style);
    const fontSizePx = (parseFloat(style.fontSize) || 16) * metrics.designScale;
    const fontSizePt = Number((fontSizePx * 0.5).toFixed(2));
    const fontFace = resolveFontFamily(style.fontFamily, this.fontMap);
    const bold = isBoldWeight(style.fontWeight);
    const italic = style.fontStyle === 'italic';
    const align = mapTextAlign(style.textAlign);
    const letterSpacingPx = parseFloat(style.letterSpacing);
    const charSpacingPt =
      !Number.isNaN(letterSpacingPx) && letterSpacingPx !== 0
        ? Number((letterSpacingPx * metrics.designScale * 0.5).toFixed(2))
        : undefined;

    const combinedOpacity =
      textGrad
        ? 1
        : (fontColor?.opacity ?? 1) * (Number.isNaN(elOpacity) ? 1 : elOpacity);
    const transparency =
      combinedOpacity < 1 ? Math.round((1 - combinedOpacity) * 100) : 0;

    for (const child of Array.from(el.childNodes)) {
      if (child.nodeType !== Node.TEXT_NODE) continue;
      const rawText = child.textContent ?? '';
      const collapsed = rawText.replace(/\s+/g, ' ').trim();
      if (!collapsed) continue;

      const text = applyTextTransform(collapsed, style.textTransform);
      const range = document.createRange();
      range.selectNodeContents(child);
      const textDomRect = range.getBoundingClientRect();
      const rawRect = this.toSlideRect(textDomRect, metrics);

      if (rawRect.wPx <= 0.5 || rawRect.hPx <= 0.5) continue;

      const isSingleLine = this.isSingleLineRange(range, style, metrics, false);
      const box = this.computeTextBoxRect(rawRect, align, isSingleLine, fontSizePx);
      const runColor = textGrad
        ? this.ext.registerTextGradient(slide, textGrad)
        : (fontColor?.hex ?? '000000');

      slide.addText(text, {
        x: this.ext.pxToIn(box.xPx),
        y: this.ext.pxToIn(box.yPx),
        w: this.ext.pxToIn(box.wPx),
        h: this.ext.pxToIn(box.hPx),
        fontFace,
        fontSize: fontSizePt,
        bold,
        italic,
        color: runColor,
        transparency,
        align,
        valign: isSingleLine ? 'middle' : 'top',
        margin: 0,
        wrap: !isSingleLine,
        fit: isSingleLine ? 'shrink' : 'none',
        charSpacing: charSpacingPt,
      });
    }
  }

  private async processVisualElement(
    el: HTMLElement | SVGSVGElement,
    slide: PptxGenJS.Slide,
    metrics: SlideTransformMetrics,
    backdropCanvas: HTMLCanvasElement | undefined
  ): Promise<{ descend: boolean; rect?: RectPx }> {
    if (el.hasAttribute('data-pptx-ignore')) return { descend: false };

    const style = window.getComputedStyle(el);
    if (style.display === 'none' || style.visibility === 'hidden') return { descend: false };

    const opacity = parseFloat(style.opacity);
    if (!Number.isNaN(opacity) && opacity <= 0) return { descend: false };

    const domRect = el.getBoundingClientRect();
    const rect = this.toSlideRect(domRect, metrics);
    if (rect.wPx <= 0.5 || rect.hPx <= 0.5) return { descend: false };

    // 0. Declarative data-pptx-chart JSON chart (line/area combo, bar/col, or donut)
    if (this.emitDeclarativeChartAttr(el, slide, rect)) {
      return { descend: el instanceof HTMLElement, rect };
    }

    // Explicit raster opt-in via data-pptx-raster
    if (el.hasAttribute('data-pptx-raster')) {
      if (el instanceof SVGSVGElement) {
        const asset = await rasterizeSvgToPng(el, rect, this.dpiScale);
        slide.addImage({
          data: asset.dataUrl,
          x: this.ext.pxToIn(asset.xPx),
          y: this.ext.pxToIn(asset.yPx),
          w: this.ext.pxToIn(asset.wPx),
          h: this.ext.pxToIn(asset.hPx),
        });
        return { descend: false };
      }
      const asset = hasBackdropBlur(style)
        ? await renderFrostedGlassElement(style, rect, backdropCanvas, this.dpiScale)
        : await renderMeshBlurElement(style, rect, this.dpiScale);
      slide.addImage({
        data: asset.dataUrl,
        x: this.ext.pxToIn(asset.xPx),
        y: this.ext.pxToIn(asset.yPx),
        w: this.ext.pxToIn(asset.wPx),
        h: this.ext.pxToIn(asset.hPx),
      });
      return { descend: false };
    }

    // 1. SVG Element -> auto-extracted native combo chart, native <a:custGeom>, or 3x DPI PNG
    if (el instanceof SVGSVGElement) {
      const extractedChart = extractAreaLineChartFromSvg(el, rect);
      if (extractedChart) {
        this.ext.addGradientLineChart(slide, extractedChart);
        return { descend: false };
      }

      const result = await renderSvgElement(el, rect, this.dpiScale);
      if (result.kind === 'vector') {
        for (const shapeOpts of result.shapes) {
          this.ext.addExtendedShape(slide, shapeOpts);
        }
      } else {
        slide.addImage({
          data: result.asset.dataUrl,
          x: this.ext.pxToIn(result.asset.xPx),
          y: this.ext.pxToIn(result.asset.yPx),
          w: this.ext.pxToIn(result.asset.wPx),
          h: this.ext.pxToIn(result.asset.hPx),
        });
      }
      return { descend: false };
    }

    // 2. HTML <table> Element -> native PowerPoint table via slide.addTable()
    if (el instanceof HTMLTableElement) {
      this.processTableElement(el, slide, metrics, rect);
      return { descend: false };
    }

    // 3. HTML <img> Element -> embed safely without triggering failing network fetches offline
    if (el instanceof HTMLImageElement && el.src) {
      if (el.src.startsWith('data:')) {
        slide.addImage({
          data: el.src,
          x: this.ext.pxToIn(rect.xPx),
          y: this.ext.pxToIn(rect.yPx),
          w: this.ext.pxToIn(rect.wPx),
          h: this.ext.pxToIn(rect.hPx),
        });
      } else if (el.complete && el.naturalWidth > 0) {
        let embedded = false;
        try {
          const canvas = document.createElement('canvas');
          canvas.width = el.naturalWidth;
          canvas.height = el.naturalHeight;
          const ctx = canvas.getContext('2d');
          if (ctx) {
            ctx.drawImage(el, 0, 0);
            const dataUrl = canvas.toDataURL('image/png');
            slide.addImage({
              data: dataUrl,
              x: this.ext.pxToIn(rect.xPx),
              y: this.ext.pxToIn(rect.yPx),
              w: this.ext.pxToIn(rect.wPx),
              h: this.ext.pxToIn(rect.hPx),
            });
            embedded = true;
          }
        } catch {
          // Cross-origin tainted canvas -> allow path only for same-origin URLs
        }

        if (!embedded) {
          try {
            const url = new URL(el.src, window.location.href);
            if (url.origin === window.location.origin) {
              slide.addImage({
                path: el.src,
                x: this.ext.pxToIn(rect.xPx),
                y: this.ext.pxToIn(rect.yPx),
                w: this.ext.pxToIn(rect.wPx),
                h: this.ext.pxToIn(rect.hPx),
              });
            }
          } catch {
            // Skip invalid/unreachable external URLs gracefully
          }
        }
      }
      return { descend: false };
    }

    const backdropCtx = backdropCanvas?.getContext('2d') ?? undefined;

    // 4. Frosted-glass elements (backdrop-filter: blur(...) or data-pptx-glass)
    const glassOverride = el.getAttribute('data-pptx-glass');
    const isGlass =
      hasBackdropBlur(style) ||
      glassOverride === 'native-translucent' ||
      glassOverride === 'raster-blur';

    if (isGlass) {
      const effectiveGlassMode: GlassMode =
        glassOverride === 'raster-blur' || glassOverride === 'native-translucent'
          ? glassOverride
          : this.glassMode;

      if (effectiveGlassMode === 'raster-blur') {
        const asset = await renderFrostedGlassElement(style, rect, backdropCanvas, this.dpiScale);
        slide.addImage({
          data: asset.dataUrl,
          x: this.ext.pxToIn(asset.xPx),
          y: this.ext.pxToIn(asset.yPx),
          w: this.ext.pxToIn(asset.wPx),
          h: this.ext.pxToIn(asset.hPx),
        });
      } else {
        this.emitBoxWithSelectiveBorders(slide, style, rect, metrics, true);
      }

      if (backdropCtx) {
        paintCssLayerToCanvas(backdropCtx, style, rect);
      }
      return { descend: true, rect };
    }

    // 5. Multi-layer mesh blur / radial-gradient / complex non-linear gradient -> 3x DPI PNG
    if (!isBackgroundClipText(style) && isComplexGradientOrMesh(style)) {
      const asset = await renderMeshBlurElement(style, rect, this.dpiScale);
      slide.addImage({
        data: asset.dataUrl,
        x: this.ext.pxToIn(asset.xPx),
        y: this.ext.pxToIn(asset.yPx),
        w: this.ext.pxToIn(asset.wPx),
        h: this.ext.pxToIn(asset.hPx),
      });
      if (backdropCtx) {
        paintCssLayerToCanvas(backdropCtx, style, rect);
      }
      return { descend: true, rect };
    }

    // 6. Standard geometry / multi-stop linear-gradient / Aurora softEdge glow blob -> native editable PPTX shape
    if (hasVisualBoxStyle(style)) {
      this.emitBoxWithSelectiveBorders(slide, style, rect, metrics, false);

      if (backdropCtx) {
        paintCssLayerToCanvas(backdropCtx, style, rect);
      }
    }

    return { descend: true, rect };
  }

  private async walkElementTree(
    el: HTMLElement | SVGSVGElement,
    slide: PptxGenJS.Slide,
    metrics: SlideTransformMetrics,
    backdropCanvas: HTMLCanvasElement | undefined
  ): Promise<void> {
    const { descend, rect } = await this.processVisualElement(
      el,
      slide,
      metrics,
      backdropCanvas
    );
    if (!descend || !(el instanceof HTMLElement)) return;

    if (rect) {
      await this.emitPseudoElement(el, '::before', slide, rect, metrics, backdropCanvas);
    }

    const isDeclarativeChartContainer = el.hasAttribute('data-pptx-chart');
    let handledRichTextChildren = false;

    if (!isDeclarativeChartContainer) {
      handledRichTextChildren = this.emitRichTextBlock(el, slide, metrics);
      if (!handledRichTextChildren) {
        this.emitDirectTextNodes(el, slide, metrics);
      }
    }

    if (!handledRichTextChildren) {
      const sortedChildren = sortChildrenInPaintOrder(el);
      for (const child of sortedChildren) {
        if (isDeclarativeChartContainer && child instanceof SVGSVGElement) {
          continue;
        }
        await this.walkElementTree(child, slide, metrics, backdropCanvas);
      }
    }

    if (rect) {
      await this.emitPseudoElement(el, '::after', slide, rect, metrics, backdropCanvas);
    }
  }

  /**
   * Walks a slide container in z-index paint order, normalizing scaled or
   * hidden (`display: none`) slides into exact 1920x1080 -> 13.333" x 7.5" coordinates.
   */
  public async convertSlide(
    slideElement: HTMLElement,
    targetSlide?: PptxGenJS.Slide
  ): Promise<PptxGenJS.Slide> {
    return this.withVisibleSlide(slideElement, async () => {
      const slide = targetSlide ?? this.ext.addSlide();
      const metrics = this.computeMetrics(slideElement);
      const needsBackdropCanvas =
        this.glassMode === 'raster-blur' ||
        Boolean(slideElement.querySelector('[data-pptx-glass="raster-blur"]'));
      const backdropCanvas = needsBackdropCanvas ? this.createBackdropCanvas() : undefined;

      await this.walkElementTree(slideElement, slide, metrics, backdropCanvas);
      return slide;
    });
  }

  /**
   * Converts multiple HTML slide containers (including hidden `display: none` or
   * CSS-scaled slides) into PowerPoint slides in order.
   */
  public async convertDeck(
    slideElements: Iterable<HTMLElement>
  ): Promise<PptxGenJS.Slide[]> {
    const slides: PptxGenJS.Slide[] = [];
    for (const el of slideElements) {
      slides.push(await this.convertSlide(el));
    }
    return slides;
  }
}
