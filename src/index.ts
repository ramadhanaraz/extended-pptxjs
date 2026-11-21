export {
  ExtendedPptxJS,
  type PptxShapeName,
  type GradientStop,
  type ExtendedGradient,
  type ExtendedShadow,
  type EmbeddedFont,
  type ExtendedPptxJSOptions,
  type ExtendedShapeOptions,
  type GradientLineChartOptions,
  type GradientBarChartOptions,
  type DonutChartOptions,
} from './core/ExtendedPptxJS';

export {
  DEFAULT_FONT_MAP,
  parseCssColor,
  parseLinearGradient,
  parseBorderRadius,
  parseBoxShadow,
  parseBorder,
  parseFilterBlur,
  parseCssToShapeOptions,
  resolveFontFamily,
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
  extractAreaLineChartFromSvg,
  type GlassMode,
  type HtmlSlideConverterOptions,
} from './walker/HtmlSlideConverter';
