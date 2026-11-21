import type PptxGenJS from 'pptxgenjs';
import { ExtendedPptxJS } from '../core/ExtendedPptxJS';
import {
  parseCssColor,
  parseCssToShapeOptions,
  parseFilterBlur,
  parseLinearGradient,
  type RectPx,
} from '../parsers/cssParser';
import {
  DEFAULT_RASTER_DPI_SCALE,
  hasBackdropBlur,
  isComplexGradientOrMesh,
  paintCssLayerToCanvas,
  renderFrostedGlassElement,
  renderMeshBlurElement,
  renderSvgElement,
} from '../renderers/hybridAssetRenderer';

export interface HtmlSlideConverterOptions {
  dpiScale?: number;
  includeText?: boolean;
  baseWidthPx?: number;
  baseHeightPx?: number;
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

function normalizeFontName(fontFamily: string): string {
  const first = fontFamily.split(',')[0]?.trim() ?? 'Arial';
  return first.replace(/^["']|["']$/g, '') || 'Arial';
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

export class HtmlSlideConverter {
  private readonly dpiScale: number;
  private readonly includeText: boolean;
  private readonly baseWidthPx: number;
  private readonly baseHeightPx: number;

  constructor(
    private readonly ext: ExtendedPptxJS,
    options: HtmlSlideConverterOptions = {}
  ) {
    this.dpiScale = options.dpiScale ?? DEFAULT_RASTER_DPI_SCALE;
    this.includeText = options.includeText ?? true;
    this.baseWidthPx = options.baseWidthPx ?? ext.baseWidthPx;
    this.baseHeightPx = options.baseHeightPx ?? ext.baseHeightPx;
  }

  private toSlideRect(
    elRect: DOMRect,
    containerRect: DOMRect
  ): RectPx {
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

    // 1. SVG Element -> native <a:custGeom> or 3x DPI transparent PNG fallback
    if (el instanceof SVGSVGElement) {
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
      return false; // Do not descend into SVG child primitives
    }

    // 2. HTML <img> Element -> embed directly when possible
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

    // 3. Frosted-glass card (backdrop-filter: blur(...)) -> 3x DPI PNG surface + keep children native!
    if (hasBackdropBlur(style)) {
      const asset = await renderFrostedGlassElement(style, rect, backdropCanvas, this.dpiScale);
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

    // 4. Multi-layer mesh blur / radial-gradient / complex non-linear gradient -> 3x DPI PNG
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

    // 5. Standard geometry / multi-stop linear-gradient / Aurora softEdge glow blob -> native editable PPTX shape
    if (hasVisualBoxStyle(style)) {
      const shapeOpts = parseCssToShapeOptions(style, rect);

      // Treat CSS filter: blur() elements without hard borders as Aurora soft-edge glow ellipses/rects
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

      // If backgroundImage has a linear-gradient, ensure shapeOpts.gradient is populated
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
    const fontFace = normalizeFontName(style.fontFamily);
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

      // Add slight width padding (+8px) to avoid premature line wrapping from font metric differences
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

    // Emit direct text nodes of this element before children in paint order
    this.emitDirectTextNodes(el, slide, containerRect);

    const sortedChildren = sortChildrenInPaintOrder(el);
    for (const child of sortedChildren) {
      await this.walkElementTree(child, slide, containerRect, backdropCanvas);
    }
  }

  /**
   * Walks a 1920x1080 HTML slide container in z-index paint order and converts it
   * into an editable PowerPoint slide with hybrid 3x DPI raster fallback only for
   * frosted glass, mesh blurs, and complex SVGs.
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
