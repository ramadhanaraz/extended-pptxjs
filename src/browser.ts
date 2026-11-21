import { ExtendedPptxJS } from './core/ExtendedPptxJS';
import * as cssParser from './parsers/cssParser';
import * as hybridAssetRenderer from './renderers/hybridAssetRenderer';
import { HtmlSlideConverter } from './walker/HtmlSlideConverter';

const globalObj = globalThis as Record<string, unknown>;

globalObj.ExtendedPptxJS = ExtendedPptxJS;
globalObj.HtmlSlideConverter = HtmlSlideConverter;
globalObj.cssParser = cssParser;
globalObj.hybridAssetRenderer = hybridAssetRenderer;

globalObj.parseCssColor = cssParser.parseCssColor;
globalObj.parseLinearGradient = cssParser.parseLinearGradient;
globalObj.parseBorderRadius = cssParser.parseBorderRadius;
globalObj.parseBoxShadow = cssParser.parseBoxShadow;
globalObj.parseBorder = cssParser.parseBorder;
globalObj.parseFilterBlur = cssParser.parseFilterBlur;
globalObj.parseCssToShapeOptions = cssParser.parseCssToShapeOptions;

export {
  ExtendedPptxJS,
  HtmlSlideConverter,
  cssParser,
  hybridAssetRenderer,
};
