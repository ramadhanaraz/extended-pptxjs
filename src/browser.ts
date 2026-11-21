import { ExtendedPptxJS } from './core/ExtendedPptxJS';
import * as cssParser from './parsers/cssParser';
import * as hybridAssetRenderer from './renderers/hybridAssetRenderer';
import { HtmlSlideConverter, extractAreaLineChartFromSvg } from './walker/HtmlSlideConverter';

const globalObj = globalThis as Record<string, unknown>;

globalObj.ExtendedPptxJS = ExtendedPptxJS;
globalObj.HtmlSlideConverter = HtmlSlideConverter;
globalObj.extractAreaLineChartFromSvg = extractAreaLineChartFromSvg;
globalObj.cssParser = cssParser;
globalObj.hybridAssetRenderer = hybridAssetRenderer;

globalObj.parseCssColor = cssParser.parseCssColor;
globalObj.parseLinearGradient = cssParser.parseLinearGradient;
globalObj.parseBorderRadius = cssParser.parseBorderRadius;
globalObj.parseBoxShadow = cssParser.parseBoxShadow;
globalObj.parseBorder = cssParser.parseBorder;
globalObj.parseFilterBlur = cssParser.parseFilterBlur;
globalObj.parseCssToShapeOptions = cssParser.parseCssToShapeOptions;
globalObj.resolveFontFamily = cssParser.resolveFontFamily;

export {
  ExtendedPptxJS,
  HtmlSlideConverter,
  extractAreaLineChartFromSvg,
  cssParser,
  hybridAssetRenderer,
};
