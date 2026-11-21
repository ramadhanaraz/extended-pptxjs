import PptxGenJS from 'pptxgenjs';
import JSZip from 'jszip';

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
  areaGradient: ExtendedGradient;
}

function escapeXmlAttr(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
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

    const sharedSeries = [
      {
        name: seriesName,
        labels: opts.labels,
        values: opts.values,
      },
    ];

    const comboTypes: PptxGenJS.IChartMulti[] = [
      {
        type: this.pptx.ChartType.area,
        data: sharedSeries,
        options: {
          chartColors: [lineColor],
        },
      },
      {
        type: this.pptx.ChartType.line,
        data: sharedSeries,
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
      areaGradient,
    });

    return chartIndex;
  }

  private buildGradFillXml(gradient: ExtendedGradient): string {
    // CSS 0deg points North (up); OOXML 0deg points East (right), clockwise
    const ooxmlDeg = (gradient.angleDeg - 90 + 360) % 360;
    const ooxmlAng = Math.round(ooxmlDeg * 60000);

    const stopsXml = gradient.stops
      .map((s) => {
        const pos = Math.round(Math.max(0, Math.min(100, s.position)) * 1000);
        const alphaXml =
          s.opacity !== undefined && s.opacity < 1
            ? `<a:alpha val="${Math.round(Math.max(0, s.opacity) * 100000)}"/>`
            : '';
        return `<a:gs pos="${pos}"><a:srgbClr val="${s.color.replace('#', '')}">${alphaXml}</a:srgbClr></a:gs>`;
      })
      .join('');

    return `<a:gradFill rotWithShape="1"><a:gsLst>${stopsXml}</a:gsLst><a:lin ang="${ooxmlAng}" scaled="1"/></a:gradFill>`;
  }

  private buildEffectLstXml(softEdgePt?: number, shadow?: ExtendedShadow): string {
    const effects: string[] = [];

    if (shadow) {
      const blurRadEmu = Math.round(shadow.blurPt * 12700);
      const distPx = Math.hypot(shadow.offsetXPx, shadow.offsetYPx);
      const distEmu = Math.round(this.pxToPt(distPx) * 12700);
      const radAngle = Math.atan2(shadow.offsetYPx, shadow.offsetXPx);
      const degAngle = ((radAngle * 180) / Math.PI + 360) % 360;
      const ooxmlDir = Math.round(degAngle * 60000);
      const alphaVal = Math.round(Math.max(0, Math.min(1, shadow.opacity ?? 0.25)) * 100000);

      effects.push(
        `<a:outerShdw blurRad="${blurRadEmu}" dist="${distEmu}" dir="${ooxmlDir}" algn="ctr" rotWithShape="0">` +
          `<a:srgbClr val="${shadow.color.replace('#', '')}"><a:alpha val="${alphaVal}"/></a:srgbClr>` +
          `</a:outerShdw>`
      );
    }

    if (softEdgePt !== undefined && softEdgePt > 0) {
      const radEmu = Math.round(softEdgePt * 12700);
      effects.push(`<a:softEdge rad="${radEmu}"/>`);
    }

    return effects.length > 0 ? `<a:effectLst>${effects.join('')}</a:effectLst>` : '';
  }

  private async applyChartPatches(zip: JSZip): Promise<void> {
    for (const patch of this.chartPatches) {
      const chartPath = `ppt/charts/chart${patch.chartIndex}.xml`;
      const file = zip.file(chartPath);
      if (!file) continue;

      let xml = await file.async('string');
      const gradXml = this.buildGradFillXml(patch.areaGradient);

      // Patch the areaChart series <c:spPr> fill with our multi-stop vertical <a:gradFill>
      xml = xml.replace(
        /(<c:areaChart>[\s\S]*?<c:ser>[\s\S]*?<c:spPr>)([\s\S]*?)(<\/c:spPr>)/,
        (_full, prefix: string, spPrInner: string, suffix: string) => {
          let updatedInner = spPrInner.replace(
            /<a:solidFill>[\s\S]*?<\/a:solidFill>|<a:solidFill\/>/,
            gradXml
          );
          if (!/<a:ln[\s>]/.test(updatedInner)) {
            updatedInner = `${updatedInner}<a:ln><a:noFill/></a:ln>`;
          }
          return `${prefix}${updatedInner}${suffix}`;
        }
      );

      // Ensure axis lines remain hidden across viewers
      xml = xml.replace(
        /(<c:(?:catAx|valAx)>[\s\S]*?<c:spPr>[\s\S]*?)(<a:ln[\s\S]*?<\/a:ln>)([\s\S]*?<\/c:(?:catAx|valAx)>)/g,
        '$1<a:ln><a:noFill/></a:ln>$3'
      );

      zip.file(chartPath, xml);
    }
  }

  private async applyEmbeddedFonts(zip: JSZip): Promise<void> {
    if (this.embeddedFonts.length === 0) return;

    // 1. Ensure Content_Types has fntdata extension
    const ctFile = zip.file('[Content_Types].xml');
    if (ctFile) {
      let ctXml = await ctFile.async('string');
      if (!ctXml.includes('Extension="fntdata"')) {
        ctXml = ctXml.replace(
          '</Types>',
          '<Default Extension="fntdata" ContentType="application/x-fontdata"/></Types>'
        );
        zip.file('[Content_Types].xml', ctXml);
      }
    }

    // 2. Write ppt/fonts/fontN.fntdata and register relationships
    const relsPath = 'ppt/_rels/presentation.xml.rels';
    const relsFile = zip.file(relsPath);
    if (!relsFile) return;

    let relsXml = await relsFile.async('string');
    const idMatches = Array.from(relsXml.matchAll(/Id="rId(\d+)"/g));
    let maxRid = idMatches.reduce((max, m) => Math.max(max, parseInt(m[1], 10)), 0);

    const embeddedFontEntries: string[] = [];
    const newRelEntries: string[] = [];

    this.embeddedFonts.forEach((font, idx) => {
      const fontFilename = `font${idx + 1}.fntdata`;
      const bytes =
        font.data instanceof Uint8Array ? font.data : new Uint8Array(font.data);
      zip.file(`ppt/fonts/${fontFilename}`, bytes);

      const rId = `rId${++maxRid}`;
      newRelEntries.push(
        `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/${fontFilename}"/>`
      );
      embeddedFontEntries.push(
        `<p:embeddedFont><p:font typeface="${escapeXmlAttr(font.name)}"/><p:regular r:id="${rId}"/></p:embeddedFont>`
      );
    });

    relsXml = relsXml.replace('</Relationships>', `${newRelEntries.join('')}</Relationships>`);
    zip.file(relsPath, relsXml);

    // 3. Inject <p:embeddedFontLst> into ppt/presentation.xml
    const presPath = 'ppt/presentation.xml';
    const presFile = zip.file(presPath);
    if (presFile) {
      let presXml = await presFile.async('string');
      const fontLstXml = `<p:embeddedFontLst>${embeddedFontEntries.join('')}</p:embeddedFontLst>`;
      if (/<p:embeddedFontLst>[\s\S]*?<\/p:embeddedFontLst>/.test(presXml)) {
        presXml = presXml.replace(
          /<p:embeddedFontLst>[\s\S]*?<\/p:embeddedFontLst>/,
          fontLstXml
        );
      } else {
        presXml = presXml.replace('</p:presentation>', `${fontLstXml}</p:presentation>`);
      }
      zip.file(presPath, presXml);
    }
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
      let xml = await file.async('string');

      const patchMap = new Map(slidePatches.map((p) => [p.objectName, p]));

      xml = xml.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, (spBlock) => {
        const nameMatch = /<p:cNvPr[^>]*name="([^"]+)"/.exec(spBlock);
        if (!nameMatch) return spBlock;
        const patch = patchMap.get(nameMatch[1]);
        if (!patch) return spBlock;

        return spBlock.replace(/<p:spPr>[\s\S]*?<\/p:spPr>/, (spPrBlock) => {
          let updatedSpPr = spPrBlock;

          if (patch.customGeometryXml) {
            updatedSpPr = updatedSpPr.replace(
              /<a:prstGeom[^>]*>[\s\S]*?<\/a:prstGeom>|<a:prstGeom[^>]*\/>/,
              patch.customGeometryXml
            );
          } else if (patch.cornerRadiusAdj !== undefined) {
            updatedSpPr = updatedSpPr.replace(
              /<a:prstGeom prst="roundRect">[\s\S]*?<\/a:prstGeom>|<a:prstGeom prst="roundRect"\/>/,
              `<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val ${patch.cornerRadiusAdj}"/></a:avLst></a:prstGeom>`
            );
          }

          if (patch.gradient) {
            const gradXml = this.buildGradFillXml(patch.gradient);
            // Replace only the first <a:solidFill> (shape fill), preserving <a:ln> border fills
            updatedSpPr = updatedSpPr.replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/, gradXml);
          }

          const effectLstXml = this.buildEffectLstXml(patch.softEdgePt, patch.shadow);
          if (effectLstXml) {
            if (/<a:effectLst[\s/>]/.test(updatedSpPr)) {
              updatedSpPr = updatedSpPr.replace(
                /<a:effectLst>[\s\S]*?<\/a:effectLst>|<a:effectLst\/>/,
                effectLstXml
              );
            } else {
              updatedSpPr = updatedSpPr.replace('</p:spPr>', `${effectLstXml}</p:spPr>`);
            }
          }

          return updatedSpPr;
        });
      });

      zip.file(slidePath, xml);
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
