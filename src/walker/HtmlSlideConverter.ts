import type PptxGenJS from 'pptxgenjs';
import {
  ExtendedPptxJS,
  type EmbeddedFont,
  type ExtendedGradient,
  type GradientLineChartOptions,
  type GradientStop,
} from '../core/ExtendedPptxJS';
import {
  parseCssColor,
  parseCssToShapeOptions,
  parseFilterBlur,
  parseLinearGradient,
  resolveFontFamily,
  type RectPx,
} from '../parsers/cssParser';
import {
  DEFAULT_RASTER_DPI_SCALE,
  hasBackdropBlur,
  isComplexGradientOrMesh,
  paintCssLayerToCanvas,
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

interface PaintItem {
  element: HTMLElement | SVGSVGElement;
  zIndex: number;
  isPositioned: boolean;
  domOrder: number;
}

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

function hasVisualBoxStyle(style: CSSStyleDeclaration): boolean {
  const bgColor = parseCssColor(style.backgroundColor);
  if (bgColor && bgColor.opacity > 0) return true;

  if (style.backgroundImage && style.backgroundImage !== 'none') return true;

  for (const side of ['Top', 'Right', 'Bottom', 'Left'] as const) {
    const w = parseFloat(style[`border${side}Width`]);
    const s = style[`border${side}Style`];
    const c = parseCssColor(style[`border${side}Color`]);
    if (!Number.isNaN(w) && w > 0 && s !== 'none' && s !== 'hidden' && c && c.opacity > 0) {
      return true;
    }
  }

  if (style.boxShadow && style.boxShadow !== 'none') return true;

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

  private toSlideRect(elRect: DOMRect, containerRect: DOMRect): RectPx {
    const scaleX = containerRect.width > 0 ? this.baseWidthPx / containerRect.width : 1;
    const scaleY = containerRect.height > 0 ? this.baseHeightPx / containerRect.height : 1;

    return {
      xPx: Number(((elRect.left - containerRect.left) * scaleX).toFixed(2)),
      yPx: Number(((elRect.top - containerRect.top) * scaleY).toFixed(2)),
      wPx: Number((elRect.width * scaleX).toFixed(2)),
      hPx: Number((elRect.height * scaleY).toFixed(2)),
    };
  }

  private createBackdropCanvas(): HTMLCanvasElement | undefined {
    if (typeof document === 'undefined') return undefined;
    const canvas = document.createElement('canvas');
    canvas.width = this.baseWidthPx;
    canvas.height = this.baseHeightPx;
    return canvas;
  }

  private parseDeclarativeChartAttr(
    el: Element,
    rect: RectPx
  ): GradientLineChartOptions | undefined {
    const raw = el.getAttribute('data-pptx-chart');
    if (!raw) return undefined;

    try {
      const parsed = JSON.parse(raw) as Partial<GradientLineChartOptions> & {
        color?: string;
        gradientStops?: GradientStop[];
      };
      if (!Array.isArray(parsed.labels) || !Array.isArray(parsed.values)) {
        return undefined;
      }

      const rawLineColor = parsed.lineColor ?? parsed.color ?? '#38BDF8';
      const parsedLineColor = parseCssColor(rawLineColor);
      const lineColor = parsedLineColor?.hex ?? rawLineColor.replace('#', '');

      let areaGradient = parsed.areaGradient;
      if (!areaGradient && Array.isArray(parsed.gradientStops)) {
        areaGradient = {
          angleDeg: 180,
          stops: parsed.gradientStops.map((s) => ({
            ...s,
            color: (parseCssColor(s.color)?.hex ?? s.color).replace('#', ''),
          })),
        };
      }

      return {
        xPx: rect.xPx,
        yPx: rect.yPx,
        wPx: rect.wPx,
        hPx: rect.hPx,
        seriesName: parsed.seriesName,
        labels: parsed.labels.map(String),
        values: parsed.values.map(Number),
        lineColor,
        lineWidthPt: parsed.lineWidthPt,
        showMarkers: parsed.showMarkers,
        markerSize: parsed.markerSize,
        markerColor: parsed.markerColor
          ? (parseCssColor(parsed.markerColor)?.hex ?? parsed.markerColor.replace('#', ''))
          : undefined,
        areaGradient,
        smooth: parsed.smooth,
        showLegend: parsed.showLegend,
        valAxisMin: parsed.valAxisMin,
        valAxisMax: parsed.valAxisMax,
      };
    } catch {
      return undefined;
    }
  }

  private processTableElement(
    tableEl: HTMLTableElement,
    slide: PptxGenJS.Slide,
    containerRect: DOMRect,
    tableRect: RectPx
  ): void {
    const rows: PptxGenJS.TableRow[] = [];
    const rowHeightsIn: number[] = [];
    let colWidthsIn: number[] | undefined;

    for (const tr of Array.from(tableEl.rows)) {
      const trRect = this.toSlideRect(tr.getBoundingClientRect(), containerRect);
      if (trRect.hPx > 0) {
        rowHeightsIn.push(this.ext.pxToIn(trRect.hPx));
      }

      const trStyle = window.getComputedStyle(tr);
      const rowCells: PptxGenJS.TableCell[] = [];
      const rowColWidths: number[] = [];

      for (const cell of Array.from(tr.cells)) {
        const cellStyle = window.getComputedStyle(cell);
        const cellRect = this.toSlideRect(cell.getBoundingClientRect(), containerRect);
        if (cell.colSpan === 1 && cellRect.wPx > 0) {
          rowColWidths.push(this.ext.pxToIn(cellRect.wPx));
        }

        const rawText = (cell.innerText || cell.textContent || '').replace(/\s+/g, ' ').trim();
        const text = applyTextTransform(rawText, cellStyle.textTransform);

        const fontColor = parseCssColor(cellStyle.color);
        const fontSizePx = parseFloat(cellStyle.fontSize) || 16;
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
              pt: Number((wPx * 0.5).toFixed(2)),
            };
          }
          return { type: 'none' };
        };

        const padT = Number(((parseFloat(cellStyle.paddingTop) || 0) * 0.5).toFixed(1));
        const padR = Number(((parseFloat(cellStyle.paddingRight) || 0) * 0.5).toFixed(1));
        const padB = Number(((parseFloat(cellStyle.paddingBottom) || 0) * 0.5).toFixed(1));
        const padL = Number(((parseFloat(cellStyle.paddingLeft) || 0) * 0.5).toFixed(1));

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

  private async processVisualElement(
    el: HTMLElement | SVGSVGElement,
    slide: PptxGenJS.Slide,
    containerRect: DOMRect,
    backdropCanvas: HTMLCanvasElement | undefined
  ): Promise<boolean> {
    if (el.hasAttribute('data-pptx-ignore')) return false;

    const style = window.getComputedStyle(el);
    if (style.display === 'none' || style.visibility === 'hidden') return false;

    const opacity = parseFloat(style.opacity);
    if (!Number.isNaN(opacity) && opacity <= 0) return false;

    const domRect = el.getBoundingClientRect();
    const rect = this.toSlideRect(domRect, containerRect);
    if (rect.wPx <= 0.5 || rect.hPx <= 0.5) return false;

    // 0. Declarative data-pptx-chart JSON combo chart
    const declarativeChart = this.parseDeclarativeChartAttr(el, rect);
    if (declarativeChart) {
      this.ext.addGradientLineChart(slide, declarativeChart);
      return false;
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
        return false;
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
      return false;
    }

    // 1. SVG Element -> auto-extracted native combo chart, native <a:custGeom>, or 3x DPI PNG
    if (el instanceof SVGSVGElement) {
      const extractedChart = extractAreaLineChartFromSvg(el, rect);
      if (extractedChart) {
        this.ext.addGradientLineChart(slide, extractedChart);
        return false;
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
      return false;
    }

    // 2. HTML <table> Element -> native PowerPoint table via slide.addTable()
    if (el instanceof HTMLTableElement) {
      this.processTableElement(el, slide, containerRect, rect);
      return false;
    }

    // 3. HTML <img> Element -> embed directly
    if (el instanceof HTMLImageElement && el.src) {
      if (el.src.startsWith('data:')) {
        slide.addImage({
          data: el.src,
          x: this.ext.pxToIn(rect.xPx),
          y: this.ext.pxToIn(rect.yPx),
          w: this.ext.pxToIn(rect.wPx),
          h: this.ext.pxToIn(rect.hPx),
        });
      } else {
        slide.addImage({
          path: el.src,
          x: this.ext.pxToIn(rect.xPx),
          y: this.ext.pxToIn(rect.yPx),
          w: this.ext.pxToIn(rect.wPx),
          h: this.ext.pxToIn(rect.hPx),
        });
      }
      return false;
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
        const shapeOpts = parseCssToShapeOptions(style, rect);
        if (!shapeOpts.shapeType) {
          shapeOpts.shapeType = 'roundRect';
        }
        if (!shapeOpts.border) {
          shapeOpts.border = {
            color: 'FFFFFF',
            widthPt: 1,
            opacity: 0.22,
          };
        }
        if (shapeOpts.fillOpacity === undefined && !shapeOpts.gradient) {
          shapeOpts.fillOpacity = 0.72;
        }
        this.ext.addExtendedShape(slide, shapeOpts);
      }

      if (backdropCtx) {
        paintCssLayerToCanvas(backdropCtx, style, rect);
      }
      return true;
    }

    // 5. Multi-layer mesh blur / radial-gradient / complex non-linear gradient -> 3x DPI PNG
    if (isComplexGradientOrMesh(style)) {
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
      return true;
    }

    // 6. Standard geometry / multi-stop linear-gradient / Aurora softEdge glow blob -> native editable PPTX shape
    if (hasVisualBoxStyle(style)) {
      const shapeOpts = parseCssToShapeOptions(style, rect);

      const blurPt = parseFilterBlur(style.filter);
      if (blurPt !== undefined && blurPt > 0) {
        shapeOpts.softEdgePt = blurPt;
        if (shapeOpts.gradient === undefined && shapeOpts.fillColor === undefined) {
          const parsedBg = parseCssColor(style.backgroundColor);
          if (parsedBg) {
            shapeOpts.fillColor = parsedBg.hex;
            shapeOpts.fillOpacity = parsedBg.opacity;
          }
        }
      }

      if (!shapeOpts.gradient && style.backgroundImage && style.backgroundImage !== 'none') {
        const grad = parseLinearGradient(style.backgroundImage);
        if (grad) shapeOpts.gradient = grad;
      }

      this.ext.addExtendedShape(slide, shapeOpts);

      if (backdropCtx) {
        paintCssLayerToCanvas(backdropCtx, style, rect);
      }
    }

    return true;
  }

  private emitDirectTextNodes(
    el: HTMLElement,
    slide: PptxGenJS.Slide,
    containerRect: DOMRect
  ): void {
    if (!this.includeText) return;

    const style = window.getComputedStyle(el);
    if (style.display === 'none' || style.visibility === 'hidden') return;

    const elOpacity = parseFloat(style.opacity);
    if (!Number.isNaN(elOpacity) && elOpacity <= 0) return;

    const fontColor = parseCssColor(style.color);
    const fontSizePx = parseFloat(style.fontSize) || 16;
    const fontSizePt = Number((fontSizePx * 0.5).toFixed(2));
    const fontFace = resolveFontFamily(style.fontFamily, this.fontMap);
    const bold = isBoldWeight(style.fontWeight);
    const italic = style.fontStyle === 'italic';
    const align = mapTextAlign(style.textAlign);
    const letterSpacingPx = parseFloat(style.letterSpacing);
    const charSpacingPt =
      !Number.isNaN(letterSpacingPx) && letterSpacingPx !== 0
        ? Number((letterSpacingPx * 0.5).toFixed(2))
        : undefined;

    const transparency =
      fontColor && fontColor.opacity * (Number.isNaN(elOpacity) ? 1 : elOpacity) < 1
        ? Math.round((1 - fontColor.opacity * (Number.isNaN(elOpacity) ? 1 : elOpacity)) * 100)
        : 0;

    for (const child of Array.from(el.childNodes)) {
      if (child.nodeType !== Node.TEXT_NODE) continue;
      const rawText = child.textContent ?? '';
      const collapsed = rawText.replace(/\s+/g, ' ').trim();
      if (!collapsed) continue;

      const text = applyTextTransform(collapsed, style.textTransform);
      const range = document.createRange();
      range.selectNodeContents(child);
      const textDomRect = range.getBoundingClientRect();
      const rect = this.toSlideRect(textDomRect, containerRect);

      if (rect.wPx <= 0.5 || rect.hPx <= 0.5) continue;

      const boxWPx = rect.wPx + 8;
      const boxXPx =
        align === 'center'
          ? rect.xPx - 4
          : align === 'right'
            ? rect.xPx - 8
            : rect.xPx;

      const isMultiLine = rect.hPx > fontSizePx * 1.7;

      slide.addText(text, {
        x: this.ext.pxToIn(boxXPx),
        y: this.ext.pxToIn(rect.yPx),
        w: this.ext.pxToIn(boxWPx),
        h: this.ext.pxToIn(Math.max(rect.hPx, fontSizePx * 1.2)),
        fontFace,
        fontSize: fontSizePt,
        bold,
        italic,
        color: fontColor?.hex ?? '000000',
        transparency,
        align,
        valign: isMultiLine ? 'top' : 'middle',
        margin: 0,
        charSpacing: charSpacingPt,
      });
    }
  }

  private async walkElementTree(
    el: HTMLElement | SVGSVGElement,
    slide: PptxGenJS.Slide,
    containerRect: DOMRect,
    backdropCanvas: HTMLCanvasElement | undefined
  ): Promise<void> {
    const shouldDescend = await this.processVisualElement(el, slide, containerRect, backdropCanvas);
    if (!shouldDescend || !(el instanceof HTMLElement)) return;

    this.emitDirectTextNodes(el, slide, containerRect);

    const sortedChildren = sortChildrenInPaintOrder(el);
    for (const child of sortedChildren) {
      await this.walkElementTree(child, slide, containerRect, backdropCanvas);
    }
  }

  /**
   * Walks a 1920x1080 HTML slide container in z-index paint order and converts it
   * into an editable PowerPoint slide using native translucent glass shapes by default,
   * native combo charts, native tables, and optional 3x DPI raster fallback.
   */
  public async convertSlide(
    slideElement: HTMLElement,
    targetSlide?: PptxGenJS.Slide
  ): Promise<PptxGenJS.Slide> {
    const slide = targetSlide ?? this.ext.addSlide();
    const containerRect = slideElement.getBoundingClientRect();
    const backdropCanvas = this.createBackdropCanvas();

    await this.walkElementTree(slideElement, slide, containerRect, backdropCanvas);
    return slide;
  }

  /**
   * Converts multiple 1920x1080 HTML slide containers into slides in order.
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
