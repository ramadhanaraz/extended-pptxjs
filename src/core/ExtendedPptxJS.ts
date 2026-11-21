import PptxGenJS from 'pptxgenjs';
import JSZip from 'jszip';
import {
  DOMParser,
  XMLSerializer,
  type Document as XmlDocument,
  type Element as XmlElement,
  type Node as XmlNode,
} from '@xmldom/xmldom';

export type PptxShapeName = Parameters<PptxGenJS.Slide['addShape']>[0];

export interface GradientStop {
  color: string;    // 6-char hex without '#', e.g. '3B82F6'
  position: number; // 0 to 100
  opacity?: number; // 0 to 1 (default 1)
}

export interface ExtendedGradient {
  angleDeg: number; // CSS angle: 0deg = top, 90deg = right, 135deg = bottom-right
  stops: GradientStop[];
}

export interface ExtendedShadow {
  color: string;     // 6-char hex without '#'
  opacity?: number;  // 0 to 1 (default 0.25)
  blurPt: number;    // Blur radius in points
  offsetXPx: number; // Horizontal offset in 1920px canvas space
  offsetYPx: number; // Vertical offset in 1920px canvas space
}

export interface EmbeddedFont {
  name: string;
  data: ArrayBuffer | Uint8Array;
}

export interface ExtendedPptxJSOptions {
  embedFonts?: EmbeddedFont[];
}

export interface ExtendedShapeOptions {
  xPx: number;
  yPx: number;
  wPx: number;
  hPx: number;
  shapeType?: PptxShapeName;
  cornerRadiusPx?: number;
  fillColor?: string;
  fillOpacity?: number;
  gradient?: ExtendedGradient;
  softEdgePt?: number; // Blurred edge radius in points (for Aurora glow blobs)
  shadow?: ExtendedShadow;
  customGeometryXml?: string; // Raw OOXML <a:custGeom>...</a:custGeom> XML
  border?: {
    color: string;
    widthPt: number;
    opacity?: number;
  };
}

export interface GradientLineChartOptions {
  xPx: number;
  yPx: number;
  wPx: number;
  hPx: number;
  seriesName?: string;
  labels: string[];
  values: number[];
  lineColor?: string;
  lineWidthPt?: number;
  showMarkers?: boolean;
  markerSize?: number;
  markerColor?: string;
  areaGradient?: ExtendedGradient;
  smooth?: boolean;
  showLegend?: boolean;
  valAxisMin?: number;
  valAxisMax?: number;
}

export interface GradientBarChartOptions {
  xPx: number;
  yPx: number;
  wPx: number;
  hPx: number;
  seriesName?: string;
  labels: string[];
  values: number[];
  barDir?: 'col' | 'bar';
  barColor?: string;
  gradient?: ExtendedGradient;
  showDataLabels?: boolean;
  dataLabelColor?: string;
  dataLabelPosition?: 'outEnd' | 'inEnd' | 'ctr' | 'bestFit';
  showGridLines?: boolean;
  showAxes?: boolean;
  showLegend?: boolean;
  barGapWidthPct?: number;
  valAxisMin?: number;
  valAxisMax?: number;
}

export interface DonutChartOptions {
  xPx: number;
  yPx: number;
  wPx: number;
  hPx: number;
  seriesName?: string;
  labels: string[];
  values: number[];
  colors?: string[];
  holeSize?: number;
  showLegend?: boolean;
  showDataLabels?: boolean;
}

interface PendingXmlPatch {
  slideIndex: number;
  objectName: string;
  gradient?: ExtendedGradient;
  softEdgePt?: number;
  shadow?: ExtendedShadow;
  customGeometryXml?: string;
  cornerRadiusAdj?: number;
}

interface PendingChartPatch {
  chartIndex: number;
  chartKind: 'combo-area-line' | 'bar';
  seriesIndex: number;
  gradient: ExtendedGradient;
  hideLineOnTargetSeries: boolean;
  hideAxes: boolean;
}

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';

const ELEMENT_NODE = 1;

// Strict OOXML CT_ShapeProperties element sequence
const SP_PR_ORDER: readonly string[] = [
  'xfrm',
  'custGeom',
  'prstGeom',
  'noFill',
  'solidFill',
  'gradFill',
  'blipFill',
  'pattFill',
  'grpFill',
  'ln',
  'effectLst',
  'effectDag',
  'scene3d',
  'sp3d',
  'extLst',
];

const FILL_TAG_NAMES = new Set([
  'noFill',
  'solidFill',
  'gradFill',
  'blipFill',
  'pattFill',
  'grpFill',
]);

function getLocalName(node: XmlNode): string {
  const el = node as XmlElement;
  if (el.localName) return el.localName;
  const name = el.tagName ?? el.nodeName ?? '';
  const colonIdx = name.indexOf(':');
  return colonIdx >= 0 ? name.slice(colonIdx + 1) : name;
}

function getDirectChildren(parent: XmlElement, localName?: string): XmlElement[] {
  const result: XmlElement[] = [];
  for (let i = 0; i < parent.childNodes.length; i++) {
    const child = parent.childNodes.item(i);
    if (child && child.nodeType === ELEMENT_NODE) {
      if (!localName || getLocalName(child) === localName) {
        result.push(child as XmlElement);
      }
    }
  }
  return result;
}

function getFirstDirectChild(parent: XmlElement, localName: string): XmlElement | null {
  return getDirectChildren(parent, localName)[0] ?? null;
}

function removeAllChildren(el: XmlElement): void {
  while (el.firstChild) {
    el.removeChild(el.firstChild);
  }
}

function sortSpPrChildren(spPr: XmlElement): void {
  const elements = getDirectChildren(spPr);
  removeAllChildren(spPr);
  elements.sort((a, b) => {
    const ia = SP_PR_ORDER.indexOf(getLocalName(a));
    const ib = SP_PR_ORDER.indexOf(getLocalName(b));
    const oa = ia === -1 ? SP_PR_ORDER.length : ia;
    const ob = ib === -1 ? SP_PR_ORDER.length : ib;
    return oa - ob;
  });
  for (const child of elements) {
    spPr.appendChild(child);
  }
}

function parseXmlDocument(xml: string): XmlDocument {
  const errors: string[] = [];
  const parser = new DOMParser({
    onError: (level, msg) => {
      if (level === 'error' || level === 'fatalError') {
        errors.push(`${level}: ${msg}`);
      }
    },
  });
  const doc = parser.parseFromString(xml, 'application/xml');
  if (errors.length > 0) {
    throw new Error(`XML parse error: ${errors.join('; ')}`);
  }
  return doc;
}

function serializeXmlDocument(doc: XmlDocument): string {
  return new XMLSerializer().serializeToString(doc);
}

export class ExtendedPptxJS {
  public readonly pptx: PptxGenJS;
  private patches: PendingXmlPatch[] = [];
  private chartPatches: PendingChartPatch[] = [];
  private embeddedFonts: EmbeddedFont[] = [];
  private slideIndexMap = new WeakMap<PptxGenJS.Slide, number>();
  private slideCount = 0;
  private shapeCounter = 0;
  private chartCounter = 0;

  // Base 1920x1080 canvas mapped to standard 16:9 widescreen (13.333" x 7.5")
  public readonly baseWidthPx = 1920;
  public readonly baseHeightPx = 1080;
  public readonly slideWidthIn = 13.333;
  public readonly slideHeightIn = 7.5;

  constructor(options: ExtendedPptxJSOptions = {}) {
    this.pptx = new PptxGenJS();
    this.pptx.defineLayout({
      name: 'WIDESCREEN_16_9',
      width: this.slideWidthIn,
      height: this.slideHeightIn,
    });
    this.pptx.layout = 'WIDESCREEN_16_9';

    if (options.embedFonts?.length) {
      this.embedFonts(options.embedFonts);
    }
  }

  public embedFont(font: EmbeddedFont): void {
    const existingIdx = this.embeddedFonts.findIndex(
      (f) => f.name.toLowerCase() === font.name.toLowerCase()
    );
    if (existingIdx >= 0) {
      this.embeddedFonts[existingIdx] = font;
    } else {
      this.embeddedFonts.push(font);
    }
  }

  public embedFonts(fonts: EmbeddedFont[]): void {
    for (const font of fonts) {
      this.embedFont(font);
    }
  }

  public pxToIn(px: number): number {
    return Number(((px / this.baseWidthPx) * this.slideWidthIn).toFixed(4));
  }

  public pxToPt(px: number): number {
    // 13.333" * 72 pt/in = 960 pt across 1920px -> 0.5 pt per px
    return Number((this.pxToIn(px) * 72).toFixed(2));
  }

  public addSlide(): PptxGenJS.Slide {
    const slide = this.pptx.addSlide();
    this.slideIndexMap.set(slide, this.slideCount++);
    return slide;
  }

  public addExtendedShape(
    slide: PptxGenJS.Slide,
    opts: ExtendedShapeOptions,
    explicitSlideIndex?: number
  ): string {
    const slideIndex = explicitSlideIndex ?? this.slideIndexMap.get(slide);
    if (slideIndex === undefined) {
      throw new Error(
        'Slide index unknown. Create slides via extendedPptx.addSlide() or pass explicitSlideIndex.'
      );
    }

    const objectName = `ext_shape_${++this.shapeCounter}`;
    const shapeType = opts.shapeType ?? this.pptx.ShapeType.roundRect;
    const baseFillColor = (opts.gradient?.stops[0]?.color ?? opts.fillColor ?? 'FFFFFF').replace('#', '');
    const baseTransparency =
      opts.fillOpacity !== undefined ? Math.round((1 - opts.fillOpacity) * 100) : 0;

    const shapeProps: PptxGenJS.ShapeProps = {
      objectName,
      x: this.pxToIn(opts.xPx),
      y: this.pxToIn(opts.yPx),
      w: this.pxToIn(opts.wPx),
      h: this.pxToIn(opts.hPx),
      fill: {
        color: baseFillColor,
        transparency: baseTransparency,
      },
      line: opts.border
        ? {
            color: opts.border.color.replace('#', ''),
            width: opts.border.widthPt,
            transparency:
              opts.border.opacity !== undefined ? Math.round((1 - opts.border.opacity) * 100) : 0,
          }
        : { type: 'none' },
    };

    let cornerRadiusAdj: number | undefined;
    if (
      opts.cornerRadiusPx !== undefined &&
      shapeType === this.pptx.ShapeType.roundRect &&
      !opts.customGeometryXml
    ) {
      const minDim = Math.min(opts.wPx, opts.hPx);
      shapeProps.rectRadius = minDim > 0 ? Math.min(1, (opts.cornerRadiusPx * 2) / minDim) : 0;
      cornerRadiusAdj =
        minDim > 0
          ? Math.round(Math.min(50000, Math.max(0, (opts.cornerRadiusPx / minDim) * 100000)))
          : 0;
    }

    slide.addShape(shapeType, shapeProps);

    if (
      opts.gradient ||
      opts.softEdgePt ||
      opts.shadow ||
      opts.customGeometryXml ||
      cornerRadiusAdj !== undefined
    ) {
      this.patches.push({
        slideIndex,
        objectName,
        gradient: opts.gradient,
        softEdgePt: opts.softEdgePt,
        shadow: opts.shadow,
        customGeometryXml: opts.customGeometryXml,
        cornerRadiusAdj,
      });
    }

    return objectName;
  }

  /**
   * Adds a native PowerPoint Combo Chart (area + line sharing identical values)
   * with hidden gridlines/axes and a vertical multi-stop `<a:gradFill>` patched
   * into `ppt/charts/chartN.xml` on the area series so chart data remains editable.
   */
  public addGradientLineChart(
    slide: PptxGenJS.Slide,
    opts: GradientLineChartOptions
  ): number {
    const lineColor = (opts.lineColor ?? '38BDF8').replace('#', '').toUpperCase();
    const markerColor = (opts.markerColor ?? lineColor).replace('#', '').toUpperCase();
    const seriesName = opts.seriesName ?? 'Series 1';
    const lineWidthPt = opts.lineWidthPt ?? 3;
    const showMarkers = opts.showMarkers ?? true;
    const markerSize = opts.markerSize ?? 7;

    const areaGradient: ExtendedGradient = opts.areaGradient ?? {
      angleDeg: 180,
      stops: [
        { color: lineColor, position: 0, opacity: 0.42 },
        { color: lineColor, position: 100, opacity: 0 },
      ],
    };

    // IMPORTANT: Distinct series object instances prevent pptxgenjs from overwriting _dataIndex
    const areaSeries = [
      {
        name: seriesName,
        labels: [...opts.labels],
        values: [...opts.values],
      },
    ];
    const lineSeries = [
      {
        name: seriesName,
        labels: [...opts.labels],
        values: [...opts.values],
      },
    ];

    const comboTypes: PptxGenJS.IChartMulti[] = [
      {
        type: this.pptx.ChartType.area,
        data: areaSeries,
        options: {
          chartColors: [lineColor],
        },
      },
      {
        type: this.pptx.ChartType.line,
        data: lineSeries,
        options: {
          chartColors: [lineColor],
          lineSize: lineWidthPt,
          lineSmooth: opts.smooth ?? false,
          lineDataSymbol: showMarkers ? 'circle' : 'none',
          lineDataSymbolSize: markerSize,
          lineDataSymbolLineColor: markerColor,
        },
      },
    ];

    const chartOpts: PptxGenJS.IChartOpts = {
      x: this.pxToIn(opts.xPx),
      y: this.pxToIn(opts.yPx),
      w: this.pxToIn(opts.wPx),
      h: this.pxToIn(opts.hPx),
      catGridLine: { style: 'none' },
      valGridLine: { style: 'none' },
      catAxisHidden: true,
      valAxisHidden: true,
      catAxisLineShow: false,
      valAxisLineShow: false,
      showLegend: opts.showLegend ?? false,
      showTitle: false,
    };

    if (opts.valAxisMin !== undefined) {
      chartOpts.valAxisMinVal = opts.valAxisMin;
    }
    if (opts.valAxisMax !== undefined) {
      chartOpts.valAxisMaxVal = opts.valAxisMax;
    }

    slide.addChart(comboTypes, chartOpts as unknown as unknown[]);

    const chartIndex = ++this.chartCounter;
    this.chartPatches.push({
      chartIndex,
      chartKind: 'combo-area-line',
      seriesIndex: 0,
      gradient: areaGradient,
      hideLineOnTargetSeries: true,
      hideAxes: true,
    });

    return chartIndex;
  }

  /**
   * Adds a native PowerPoint Column / Bar Chart with a multi-stop `<a:gradFill>`
   * patched onto its `<c:barChart>` series in `ppt/charts/chartN.xml`.
   */
  public addGradientBarChart(
    slide: PptxGenJS.Slide,
    opts: GradientBarChartOptions
  ): number {
    const barColor = (opts.barColor ?? '6366F1').replace('#', '').toUpperCase();
    const seriesName = opts.seriesName ?? 'Series 1';
    const showAxes = opts.showAxes ?? false;
    const showGridLines = opts.showGridLines ?? false;

    const gradient: ExtendedGradient = opts.gradient ?? {
      angleDeg: opts.barDir === 'bar' ? 90 : 180,
      stops: [
        { color: barColor, position: 0, opacity: 1 },
        { color: barColor, position: 100, opacity: 0.45 },
      ],
    };

    const chartOpts: PptxGenJS.IChartOpts = {
      x: this.pxToIn(opts.xPx),
      y: this.pxToIn(opts.yPx),
      w: this.pxToIn(opts.wPx),
      h: this.pxToIn(opts.hPx),
      barDir: opts.barDir ?? 'col',
      barGapWidthPct: opts.barGapWidthPct ?? 65,
      chartColors: [barColor],
      showValue: opts.showDataLabels ?? true,
      dataLabelPosition: opts.dataLabelPosition ?? 'outEnd',
      dataLabelColor: (opts.dataLabelColor ?? '0F172A').replace('#', '').toUpperCase(),
      catGridLine: showGridLines ? { color: 'CBD5E1', style: 'dash', size: 1 } : { style: 'none' },
      valGridLine: showGridLines ? { color: 'CBD5E1', style: 'dash', size: 1 } : { style: 'none' },
      catAxisHidden: !showAxes,
      valAxisHidden: !showAxes,
      catAxisLineShow: showAxes,
      valAxisLineShow: showAxes,
      showLegend: opts.showLegend ?? false,
      showTitle: false,
    };

    if (opts.valAxisMin !== undefined) {
      chartOpts.valAxisMinVal = opts.valAxisMin;
    }
    if (opts.valAxisMax !== undefined) {
      chartOpts.valAxisMaxVal = opts.valAxisMax;
    }

    slide.addChart(
      this.pptx.ChartType.bar,
      [{ name: seriesName, labels: [...opts.labels], values: [...opts.values] }],
      chartOpts
    );

    const chartIndex = ++this.chartCounter;
    this.chartPatches.push({
      chartIndex,
      chartKind: 'bar',
      seriesIndex: 0,
      gradient,
      hideLineOnTargetSeries: true,
      hideAxes: !showAxes,
    });

    return chartIndex;
  }

  /**
   * Adds a native PowerPoint Doughnut Chart with configurable hole size and slice colors.
   */
  public addDonutChart(
    slide: PptxGenJS.Slide,
    opts: DonutChartOptions
  ): number {
    const colors = (opts.colors?.length ? opts.colors : ['6366F1', 'E2E8F0']).map((c) =>
      c.replace('#', '').toUpperCase()
    );

    slide.addChart(
      this.pptx.ChartType.doughnut,
      [
        {
          name: opts.seriesName ?? 'Breakdown',
          labels: [...opts.labels],
          values: [...opts.values],
        },
      ],
      {
        x: this.pxToIn(opts.xPx),
        y: this.pxToIn(opts.yPx),
        w: this.pxToIn(opts.wPx),
        h: this.pxToIn(opts.hPx),
        holeSize: opts.holeSize ?? 68,
        chartColors: colors,
        showLegend: opts.showLegend ?? false,
        showValue: opts.showDataLabels ?? false,
        showPercent: opts.showDataLabels ?? false,
        showTitle: false,
      }
    );

    return ++this.chartCounter;
  }

  private buildGradFillElement(doc: XmlDocument, gradient: ExtendedGradient): XmlElement {
    const ooxmlDeg = (gradient.angleDeg - 90 + 360) % 360;
    const ooxmlAng = Math.round(ooxmlDeg * 60000);

    const gradFill = doc.createElementNS(A_NS, 'a:gradFill');
    gradFill.setAttribute('rotWithShape', '1');

    const gsLst = doc.createElementNS(A_NS, 'a:gsLst');
    for (const s of gradient.stops) {
      const pos = Math.round(Math.max(0, Math.min(100, s.position)) * 1000);
      const gs = doc.createElementNS(A_NS, 'a:gs');
      gs.setAttribute('pos', String(pos));

      const srgbClr = doc.createElementNS(A_NS, 'a:srgbClr');
      srgbClr.setAttribute('val', s.color.replace('#', '').toUpperCase());

      if (s.opacity !== undefined && s.opacity < 1) {
        const alpha = doc.createElementNS(A_NS, 'a:alpha');
        alpha.setAttribute('val', String(Math.round(Math.max(0, s.opacity) * 100000)));
        srgbClr.appendChild(alpha);
      }

      gs.appendChild(srgbClr);
      gsLst.appendChild(gs);
    }

    const lin = doc.createElementNS(A_NS, 'a:lin');
    lin.setAttribute('ang', String(ooxmlAng));
    lin.setAttribute('scaled', '1');

    gradFill.appendChild(gsLst);
    gradFill.appendChild(lin);
    return gradFill;
  }

  private buildEffectLstElement(
    doc: XmlDocument,
    softEdgePt?: number,
    shadow?: ExtendedShadow
  ): XmlElement | null {
    const hasShadow = Boolean(shadow);
    const hasSoftEdge = softEdgePt !== undefined && softEdgePt > 0;
    if (!hasShadow && !hasSoftEdge) return null;

    const effectLst = doc.createElementNS(A_NS, 'a:effectLst');

    if (shadow) {
      const blurRadEmu = Math.round(shadow.blurPt * 12700);
      const distPx = Math.hypot(shadow.offsetXPx, shadow.offsetYPx);
      const distEmu = Math.round(this.pxToPt(distPx) * 12700);
      const radAngle = Math.atan2(shadow.offsetYPx, shadow.offsetXPx);
      const degAngle = ((radAngle * 180) / Math.PI + 360) % 360;
      const ooxmlDir = Math.round(degAngle * 60000);
      const alphaVal = Math.round(Math.max(0, Math.min(1, shadow.opacity ?? 0.25)) * 100000);

      const outerShdw = doc.createElementNS(A_NS, 'a:outerShdw');
      outerShdw.setAttribute('blurRad', String(blurRadEmu));
      outerShdw.setAttribute('dist', String(distEmu));
      outerShdw.setAttribute('dir', String(ooxmlDir));
      outerShdw.setAttribute('algn', 'ctr');
      outerShdw.setAttribute('rotWithShape', '0');

      const srgbClr = doc.createElementNS(A_NS, 'a:srgbClr');
      srgbClr.setAttribute('val', shadow.color.replace('#', '').toUpperCase());

      const alpha = doc.createElementNS(A_NS, 'a:alpha');
      alpha.setAttribute('val', String(alphaVal));

      srgbClr.appendChild(alpha);
      outerShdw.appendChild(srgbClr);
      effectLst.appendChild(outerShdw);
    }

    if (softEdgePt !== undefined && softEdgePt > 0) {
      const softEdge = doc.createElementNS(A_NS, 'a:softEdge');
      softEdge.setAttribute('rad', String(Math.round(softEdgePt * 12700)));
      effectLst.appendChild(softEdge);
    }

    return effectLst;
  }

  private replaceSpPrFill(doc: XmlDocument, spPr: XmlElement, fillEl: XmlElement): void {
    for (const child of getDirectChildren(spPr)) {
      if (FILL_TAG_NAMES.has(getLocalName(child))) {
        spPr.removeChild(child);
      }
    }
    spPr.appendChild(fillEl);
  }

  private applyChartPatches(zip: JSZip): Promise<void> {
    return (async () => {
      for (const patch of this.chartPatches) {
        const chartPath = `ppt/charts/chart${patch.chartIndex}.xml`;
        const file = zip.file(chartPath);
        if (!file) continue;

        const doc = parseXmlDocument(await file.async('string'));
        const plotArea = doc.getElementsByTagNameNS(C_NS, 'plotArea')[0];
        if (!plotArea) continue;

        // Collect defined axis IDs inside plotArea to remove any orphan axId references
        const validAxisIds = new Set<string>();
        for (const axTag of ['catAx', 'valAx', 'dateAx', 'serAx']) {
          for (const ax of getDirectChildren(plotArea, axTag)) {
            const axIdEl = getFirstDirectChild(ax, 'axId');
            const val = axIdEl?.getAttribute('val');
            if (val) validAxisIds.add(val);

            if (patch.hideAxes && (axTag === 'catAx' || axTag === 'valAx')) {
              let axSpPr = getFirstDirectChild(ax, 'spPr');
              if (!axSpPr) {
                axSpPr = doc.createElementNS(C_NS, 'c:spPr');
                ax.appendChild(axSpPr);
              }
              for (const ln of getDirectChildren(axSpPr, 'ln')) {
                axSpPr.removeChild(ln);
              }
              const ln = doc.createElementNS(A_NS, 'a:ln');
              ln.appendChild(doc.createElementNS(A_NS, 'a:noFill'));
              axSpPr.appendChild(ln);
              sortSpPrChildren(axSpPr);
            }
          }
        }

        const targetChartTag = patch.chartKind === 'combo-area-line' ? 'areaChart' : 'barChart';
        const chartContainers = getDirectChildren(plotArea, targetChartTag);
        const targetChartEl = chartContainers[0];

        if (targetChartEl) {
          const seriesList = getDirectChildren(targetChartEl, 'ser');
          const targetSer = seriesList[patch.seriesIndex] ?? seriesList[0];
          if (targetSer) {
            let spPr = getFirstDirectChild(targetSer, 'spPr');
            if (!spPr) {
              spPr = doc.createElementNS(C_NS, 'c:spPr');
              targetSer.appendChild(spPr);
            }

            const gradEl = this.buildGradFillElement(doc, patch.gradient);
            this.replaceSpPrFill(doc, spPr, gradEl);

            if (patch.hideLineOnTargetSeries) {
              for (const ln of getDirectChildren(spPr, 'ln')) {
                spPr.removeChild(ln);
              }
              const ln = doc.createElementNS(A_NS, 'a:ln');
              ln.appendChild(doc.createElementNS(A_NS, 'a:noFill'));
              spPr.appendChild(ln);
            }

            sortSpPrChildren(spPr);
          }
        }

        // Prune orphan axId references across all chart types in plotArea
        if (validAxisIds.size > 0) {
          for (const chartTypeTag of ['areaChart', 'lineChart', 'barChart']) {
            for (const subChart of getDirectChildren(plotArea, chartTypeTag)) {
              for (const axIdEl of getDirectChildren(subChart, 'axId')) {
                const val = axIdEl.getAttribute('val');
                if (val && !validAxisIds.has(val)) {
                  subChart.removeChild(axIdEl);
                }
              }
            }
          }
        }

        zip.file(chartPath, serializeXmlDocument(doc));
      }
    })();
  }

  private async applyEmbeddedFonts(zip: JSZip): Promise<void> {
    if (this.embeddedFonts.length === 0) return;

    const ctFile = zip.file('[Content_Types].xml');
    if (ctFile) {
      const ctDoc = parseXmlDocument(await ctFile.async('string'));
      const typesEl = ctDoc.documentElement;
      if (typesEl) {
        const defaults = getDirectChildren(typesEl, 'Default');
        const hasFntData = defaults.some((d) => d.getAttribute('Extension') === 'fntdata');
        if (!hasFntData) {
          const defEl = ctDoc.createElementNS(CT_NS, 'Default');
          defEl.setAttribute('Extension', 'fntdata');
          defEl.setAttribute('ContentType', 'application/x-fontdata');
          typesEl.appendChild(defEl);
          zip.file('[Content_Types].xml', serializeXmlDocument(ctDoc));
        }
      }
    }

    const relsPath = 'ppt/_rels/presentation.xml.rels';
    const relsFile = zip.file(relsPath);
    if (!relsFile) return;

    const relsDoc = parseXmlDocument(await relsFile.async('string'));
    const relsRoot = relsDoc.documentElement;
    if (!relsRoot) return;

    let maxRid = 0;
    for (const rel of getDirectChildren(relsRoot, 'Relationship')) {
      const idAttr = rel.getAttribute('Id') ?? '';
      const m = /^rId(\d+)$/.exec(idAttr);
      if (m) maxRid = Math.max(maxRid, parseInt(m[1], 10));
    }

    const presPath = 'ppt/presentation.xml';
    const presFile = zip.file(presPath);
    if (!presFile) return;
    const presDoc = parseXmlDocument(await presFile.async('string'));
    const presRoot = presDoc.documentElement;
    if (!presRoot) return;

    for (const existing of getDirectChildren(presRoot, 'embeddedFontLst')) {
      presRoot.removeChild(existing);
    }
    const fontLstEl = presDoc.createElementNS(P_NS, 'p:embeddedFontLst');

    this.embeddedFonts.forEach((font, idx) => {
      const fontFilename = `font${idx + 1}.fntdata`;
      const bytes = font.data instanceof Uint8Array ? font.data : new Uint8Array(font.data);
      zip.file(`ppt/fonts/${fontFilename}`, bytes);

      const rId = `rId${++maxRid}`;
      const relEl = relsDoc.createElementNS(RELS_NS, 'Relationship');
      relEl.setAttribute('Id', rId);
      relEl.setAttribute(
        'Type',
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font'
      );
      relEl.setAttribute('Target', `fonts/${fontFilename}`);
      relsRoot.appendChild(relEl);

      const embeddedFontEl = presDoc.createElementNS(P_NS, 'p:embeddedFont');
      const fontEl = presDoc.createElementNS(P_NS, 'p:font');
      fontEl.setAttribute('typeface', font.name);
      const regularEl = presDoc.createElementNS(P_NS, 'p:regular');
      regularEl.setAttributeNS(R_NS, 'r:id', rId);

      embeddedFontEl.appendChild(fontEl);
      embeddedFontEl.appendChild(regularEl);
      fontLstEl.appendChild(embeddedFontEl);
    });

    presRoot.appendChild(fontLstEl);
    zip.file(relsPath, serializeXmlDocument(relsDoc));
    zip.file(presPath, serializeXmlDocument(presDoc));
  }

  public async exportBuffer(): Promise<ArrayBuffer> {
    const rawBuffer = (await this.pptx.write({ outputType: 'arraybuffer' })) as ArrayBuffer;
    if (
      this.patches.length === 0 &&
      this.chartPatches.length === 0 &&
      this.embeddedFonts.length === 0
    ) {
      return rawBuffer;
    }

    const zip = await JSZip.loadAsync(rawBuffer);

    const bySlide = new Map<string, PendingXmlPatch[]>();
    for (const patch of this.patches) {
      const path = `ppt/slides/slide${patch.slideIndex + 1}.xml`;
      const list = bySlide.get(path) ?? [];
      list.push(patch);
      bySlide.set(path, list);
    }

    for (const [slidePath, slidePatches] of bySlide.entries()) {
      const file = zip.file(slidePath);
      if (!file) continue;

      const doc = parseXmlDocument(await file.async('string'));
      const patchMap = new Map(slidePatches.map((p) => [p.objectName, p]));
      const spList = doc.getElementsByTagNameNS(P_NS, 'sp');

      for (let i = 0; i < spList.length; i++) {
        const sp = spList.item(i);
        if (!sp) continue;

        const nvSpPr = getFirstDirectChild(sp, 'nvSpPr');
        const cNvPr = nvSpPr ? getFirstDirectChild(nvSpPr, 'cNvPr') : null;
        const objectName = cNvPr?.getAttribute('name');
        if (!objectName) continue;

        const patch = patchMap.get(objectName);
        if (!patch) continue;

        const spPr = getFirstDirectChild(sp, 'spPr');
        if (!spPr) continue;

        if (patch.customGeometryXml) {
          const wrappedXml = patch.customGeometryXml.includes('xmlns:a=')
            ? patch.customGeometryXml
            : patch.customGeometryXml.replace(
                '<a:custGeom',
                `<a:custGeom xmlns:a="${A_NS}"`
              );
          const geomDoc = parseXmlDocument(wrappedXml);
          if (geomDoc.documentElement) {
            const importedGeom = doc.importNode(geomDoc.documentElement, true) as XmlElement;
            importedGeom.removeAttribute('xmlns:a');

            for (const geomTag of ['prstGeom', 'custGeom']) {
              for (const oldGeom of getDirectChildren(spPr, geomTag)) {
                spPr.removeChild(oldGeom);
              }
            }
            spPr.appendChild(importedGeom);
          }
        } else if (patch.cornerRadiusAdj !== undefined) {
          const prstGeom = getFirstDirectChild(spPr, 'prstGeom');
          if (prstGeom && prstGeom.getAttribute('prst') === 'roundRect') {
            for (const oldAvLst of getDirectChildren(prstGeom, 'avLst')) {
              prstGeom.removeChild(oldAvLst);
            }
            const avLst = doc.createElementNS(A_NS, 'a:avLst');
            const gd = doc.createElementNS(A_NS, 'a:gd');
            gd.setAttribute('name', 'adj');
            gd.setAttribute('fmla', `val ${patch.cornerRadiusAdj}`);
            avLst.appendChild(gd);
            prstGeom.appendChild(avLst);
          }
        }

        if (patch.gradient) {
          const gradEl = this.buildGradFillElement(doc, patch.gradient);
          this.replaceSpPrFill(doc, spPr, gradEl);
        }

        const effectLst = this.buildEffectLstElement(doc, patch.softEdgePt, patch.shadow);
        if (effectLst) {
          for (const oldEff of getDirectChildren(spPr, 'effectLst')) {
            spPr.removeChild(oldEff);
          }
          spPr.appendChild(effectLst);
        }

        sortSpPrChildren(spPr);
      }

      zip.file(slidePath, serializeXmlDocument(doc));
    }

    await this.applyChartPatches(zip);
    await this.applyEmbeddedFonts(zip);

    return await zip.generateAsync({ type: 'arraybuffer' });
  }

  public async downloadBrowserFile(fileName = 'presentation.pptx'): Promise<void> {
    const buffer = await this.exportBuffer();
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName.endsWith('.pptx') ? fileName : `${fileName}.pptx`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  }
}
