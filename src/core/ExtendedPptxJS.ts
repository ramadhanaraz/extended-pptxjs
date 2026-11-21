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
  border?: {
    color: string;
    widthPt: number;
    opacity?: number;
  };
}

interface PendingXmlPatch {
  slideIndex: number;
  objectName: string;
  gradient?: ExtendedGradient;
  softEdgePt?: number;
  shadow?: ExtendedShadow;
}

export class ExtendedPptxJS {
  public readonly pptx: PptxGenJS;
  private patches: PendingXmlPatch[] = [];
  private slideIndexMap = new WeakMap<PptxGenJS.Slide, number>();
  private slideCount = 0;
  private shapeCounter = 0;

  // Base 1920x1080 canvas mapped to standard 16:9 widescreen (13.333" x 7.5")
  public readonly baseWidthPx = 1920;
  public readonly baseHeightPx = 1080;
  public readonly slideWidthIn = 13.333;
  public readonly slideHeightIn = 7.5;

  constructor() {
    this.pptx = new PptxGenJS();
    this.pptx.defineLayout({
      name: 'WIDESCREEN_16_9',
      width: this.slideWidthIn,
      height: this.slideHeightIn,
    });
    this.pptx.layout = 'WIDESCREEN_16_9';
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

    if (opts.cornerRadiusPx !== undefined && shapeType === this.pptx.ShapeType.roundRect) {
      const minDim = Math.min(opts.wPx, opts.hPx);
      shapeProps.rectRadius = minDim > 0 ? Math.min(1, (opts.cornerRadiusPx * 2) / minDim) : 0;
    }

    slide.addShape(shapeType, shapeProps);

    if (opts.gradient || opts.softEdgePt || opts.shadow) {
      this.patches.push({
        slideIndex,
        objectName,
        gradient: opts.gradient,
        softEdgePt: opts.softEdgePt,
        shadow: opts.shadow,
      });
    }

    return objectName;
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

  public async exportBuffer(): Promise<ArrayBuffer> {
    const rawBuffer = (await this.pptx.write({ outputType: 'arraybuffer' })) as ArrayBuffer;
    if (this.patches.length === 0) return rawBuffer;

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

      for (const patch of slidePatches) {
        const spRegex = new RegExp(
          `(<p:sp>[\\s\\S]*?<p:cNvPr[^>]*name="${patch.objectName}"[\\s\\S]*?<\\/p:sp>)`
        );

        xml = xml.replace(spRegex, (spBlock) => {
          return spBlock.replace(/<p:spPr>[\s\S]*?<\/p:spPr>/, (spPrBlock) => {
            let updatedSpPr = spPrBlock;

            if (patch.gradient) {
              const gradXml = this.buildGradFillXml(patch.gradient);
              // Replace only the first <a:solidFill> (shape fill), preserving <a:ln> border fills
              updatedSpPr = updatedSpPr.replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/, gradXml);
            }

            const effectLstXml = this.buildEffectLstXml(patch.softEdgePt, patch.shadow);
            if (effectLstXml) {
              updatedSpPr = updatedSpPr.replace('</p:spPr>', `${effectLstXml}</p:spPr>`);
            }

            return updatedSpPr;
          });
        });
      }

      zip.file(slidePath, xml);
    }

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
