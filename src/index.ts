export {
  ExtendedPptxJS,
  type PptxShapeName,
  type GradientStop,
  type ExtendedGradient,
  type ExtendedShadow,
  type ExtendedShapeOptions,
} from './core/ExtendedPptxJS';

export {
  parseCssColor,
  parseLinearGradient,
  parseBorderRadius,
  parseBoxShadow,
  parseBorder,
  parseFilterBlur,
  parseCssToShapeOptions,
  splitTopLevel,
  type ParsedColor,
  type RectPx,
  type CssStyleInput,
} from './parsers/cssParser';

export {
  DEFAULT_RASTER_DPI_SCALE,
  hasBackdropBlur,
  isComplexGradientOrMesh,
  isComplexSvg,
  convertSimpleSvgToShapes,
  rasterizeSvgToPng,
  renderSvgElement,
  renderFrostedGlassElement,
  renderMeshBlurElement,
  paintCssLayerToCanvas,
  type RenderedAsset,
  type SvgRenderResult,
} from './renderers/hybridAssetRenderer';

export {
  HtmlSlideConverter,
  type HtmlSlideConverterOptions,
} from './walker/HtmlSlideConverter';
