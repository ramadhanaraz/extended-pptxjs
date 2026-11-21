import { ExtendedPptxJS } from './core/ExtendedPptxJS';
import * as cssParser from './parsers/cssParser';
import * as hybridAssetRenderer from './renderers/hybridAssetRenderer';
import { HtmlSlideConverter, extractAreaLineChartFromSvg } from './walker/HtmlSlideConverter';

function attachGlobals(target: Record<string, unknown>): void {
  target.ExtendedPptxJS = ExtendedPptxJS;
  target.HtmlSlideConverter = HtmlSlideConverter;
  target.extractAreaLineChartFromSvg = extractAreaLineChartFromSvg;
  target.cssParser = cssParser;
  target.hybridAssetRenderer = hybridAssetRenderer;

  target.parseCssColor = cssParser.parseCssColor;
  target.parseLinearGradient = cssParser.parseLinearGradient;
  target.parseBorderRadius = cssParser.parseBorderRadius;
  target.parseBoxShadow = cssParser.parseBoxShadow;
  target.parseBorder = cssParser.parseBorder;
  target.parseFilterBlur = cssParser.parseFilterBlur;
  target.parseCssToShapeOptions = cssParser.parseCssToShapeOptions;
  target.resolveFontFamily = cssParser.resolveFontFamily;
}

attachGlobals(globalThis as Record<string, unknown>);
if (typeof window !== 'undefined' && (window as unknown) !== globalThis) {
  attachGlobals(window as unknown as Record<string, unknown>);
}

export {
  ExtendedPptxJS,
  HtmlSlideConverter,
  extractAreaLineChartFromSvg,
  cssParser,
  hybridAssetRenderer,
};
