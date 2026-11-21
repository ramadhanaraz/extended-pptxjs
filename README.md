# extended-pptxjs

**OpenXML post-processor and HTML-to-PPTX engine extending [`pptxgenjs`](https://gitbrent.github.io/PptxGenJS/).**

`extended-pptxjs` converts `1920×1080` HTML/CSS presentation decks into native, fully editable `13.333" × 7.5"` (16:9 widescreen) `.pptx` files—keeping typography, tables, charts, gradients, shadows, corner radii, and simple SVGs as editable PowerPoint objects rather than flattened screenshots.

---

## Why an OpenXML Post-Processor Wrapper?

While `pptxgenjs` provides a solid programmatic slide builder, several modern CSS and OpenXML features cannot be expressed directly through its standard API:

- **Multi-stop linear gradients with per-stop alpha (`<a:gradFill>`):** `pptxgenjs` only exposes solid fills on shapes and charts. `extended-pptxjs` tags shapes via `objectName` and patches native `<a:gradFill>` XML inside `ppt/slides/slideN.xml` and `ppt/charts/chartN.xml`.
- **Exact CSS `border-radius` on pills and cards (`<a:prstGeom prst="roundRect">`):** `pptxgenjs` scales `rectRadius` inconsistently on wide pills. `extended-pptxjs` patches `<a:gd name="adj" fmla="val ..."/>` directly using `Math.round(Math.min(50000, (cornerRadiusPx / Math.min(wPx, hPx)) * 100000))`.
- **Aurora glow blobs & crisp drop shadows:** Injects native `<a:softEdge>` and `<a:outerShdw>` inside `<a:effectLst>`.
- **Editable Gradient Area + Line Combo Charts:** Pairs `<c:areaChart>` and `<c:lineChart>` sharing identical workbook series data, hides axes/gridlines, and patches a vertical multi-stop `<a:gradFill>` on the area series.
- **Native SVG vectors & Hybrid 3x DPI fallback:** Converts simple geometric SVGs into native `<a:custGeom>` paths, while optionally rasterizing complex SVGs, radial mesh blurs, and `raster-blur` glass surfaces at 3× DPI (`5760×3240` slide equivalent) via offscreen HTML5 Canvas.

---

## Quick Start

### 1. Browser / Offline Single-File HTML Decks (Zero Dependencies)

`dist/extended-pptxjs.bundle.js` is a self-contained IIFE bundle with `pptxgenjs` and `jszip` inlined—ideal for air-gapped, offline, or standalone HTML slide decks:

```html
<!-- Local copy or auto-published GitHub Pages bundle -->
<script src="https://ramadhanaraz.github.io/extended-pptxjs/dist/extended-pptxjs.bundle.js"></script>
<script>
  async function exportDeck() {
    const pptx = new globalThis.ExtendedPptxJS();
    const converter = new globalThis.HtmlSlideConverter(pptx, {
      glassMode: 'native-translucent',
      fontMap: {
        'Plus Jakarta Sans': 'Calibri',
      },
    });

    const slideEl = document.querySelector('.slide-container');
    await converter.convertSlide(slideEl);
    await pptx.downloadBrowserFile('presentation.pptx');
  }
</script>
```

The bundle attaches `ExtendedPptxJS`, `HtmlSlideConverter`, `extractAreaLineChartFromSvg`, `cssParser`, `hybridAssetRenderer`, and helper functions (`parseCssColor`, `parseLinearGradient`, `parseBorderRadius`, `parseBoxShadow`, `parseCssToShapeOptions`, `resolveFontFamily`) to `globalThis`.

### 2. ES Module (Node / Vite / TypeScript)

```bash
npm install extended-pptxjs
```

```ts
import {
  ExtendedPptxJS,
  HtmlSlideConverter,
  parseCssToShapeOptions,
} from 'extended-pptxjs';

const pptx = new ExtendedPptxJS();
const slide = pptx.addSlide();

// Convert a CSS-styled card into a native editable shape
const cardOpts = parseCssToShapeOptions(
  {
    backgroundImage: 'linear-gradient(135deg, #312e81 0%, #6d28d9 52%, #db2777 100%)',
    borderRadius: '24px',
    boxShadow: '0 20px 40px rgba(0, 0, 0, 0.45)',
    border: '1.5px solid rgba(255, 255, 255, 0.25)',
  },
  { xPx: 120, yPx: 160, wPx: 540, hPx: 320 }
);

pptx.addExtendedShape(slide, cardOpts);

const buffer = await pptx.exportBuffer();
```

---

## API Examples

### `ExtendedPptxJS`

#### `addExtendedShape(slide, opts, explicitSlideIndex?)`
Adds a native PowerPoint shape in `1920×1080` pixel coordinates and queues OpenXML patches for gradients, shadows, soft-edge blurs, custom `<a:custGeom>`, and exact `roundRect` corner radii.

```ts
const pptx = new ExtendedPptxJS();
const slide = pptx.addSlide();

// Aurora soft-edge glow blob
pptx.addExtendedShape(slide, {
  xPx: 1300,
  yPx: -120,
  wPx: 680,
  hPx: 680,
  shapeType: 'ellipse',
  fillColor: '8B5CF6',
  fillOpacity: 0.42,
  softEdgePt: 45,
});
```

#### `addGradientLineChart(slide, opts)`
Creates a natively data-editable PowerPoint Combo Chart (`area` + `line` sharing identical values) with hidden axes/gridlines, configurable stroke thickness and markers, and a vertical multi-stop `<a:gradFill>` patched into `ppt/charts/chartN.xml`.

```ts
pptx.addGradientLineChart(slide, {
  xPx: 120,
  yPx: 480,
  wPx: 760,
  hPx: 340,
  seriesName: 'Throughput (req/s)',
  labels: ['JAN', 'MAR', 'MAY', 'JUL', 'SEP', 'DEC'],
  values: [24, 48, 41, 78, 66, 94],
  lineColor: '#34D399',
  lineWidthPt: 3.5,
  showMarkers: true,
  markerSize: 7,
  areaGradient: {
    angleDeg: 180,
    stops: [
      { color: '34D399', position: 0, opacity: 0.45 },
      { color: '34D399', position: 100, opacity: 0 },
    ],
  },
});
```

#### `embedFont` / `exportBuffer` / `downloadBrowserFile`

```ts
const fontBytes = await fetch('/fonts/PlusJakartaSans-Bold.ttf').then((r) => r.arrayBuffer());

const pptx = new ExtendedPptxJS({
  embedFonts: [{ name: 'Plus Jakarta Sans', data: fontBytes }],
});

// Node / server export:
const arrayBuffer = await pptx.exportBuffer();

// Browser download:
await pptx.downloadBrowserFile('deck.pptx');
```

---

### `HtmlSlideConverter`

Walks a `1920×1080` HTML slide element using `getBoundingClientRect()` and `getComputedStyle()` in CSS `z-index` paint order.

```ts
const converter = new HtmlSlideConverter(pptx, {
  // Default: 'native-translucent' keeps backdrop-filter cards as editable translucent
  // shapes so background Aurora glows remain native & visible behind them.
  // Use 'raster-blur' to bake frosted-glass blur onto a 3x DPI transparent PNG.
  glassMode: 'native-translucent',

  // Map CSS font families (monospace stacks automatically map to Consolas by default)
  fontMap: {
    'Space Grotesk': 'Arial',
  },

  // Optional TrueType font embedding into ppt/fonts/fontN.fntdata
  embedFonts: [{ name: 'Plus Jakarta Sans', data: fontBytes }],

  // DPI multiplier for raster fallbacks (3 => 5760x3240 equivalent)
  dpiScale: 3,
});

await converter.convertSlide(document.getElementById('slide-canvas')!);
```

---

## Supported `data-pptx-*` HTML Attributes

| Attribute | Target | Behavior |
|---|---|---|
| `data-pptx-chart` | Any container or `<svg>` | Replaces the element with a native editable `addGradientLineChart()` combo chart parsed from the JSON config (`labels`, `values`, `lineColor`, `lineWidthPt`, `seriesName`, `showMarkers`, `markerSize`, `areaGradient`, `smooth`). |
| `data-pptx-glass` | Frosted-glass card | Overrides `glassMode` for a single element (`"native-translucent"` or `"raster-blur"`). |
| `data-pptx-ignore` | Any element | Skips the element and all its descendants during PPTX conversion (useful for interactive UI controls inside a slide). |
| `data-pptx-raster` | Any element or `<svg>` | Forces 3× DPI offscreen HTML5 Canvas PNG rasterization instead of native vector conversion. |

---

## Development

```bash
# Install dependencies
npm install

# TypeScript strict check
npm run typecheck

# Run Vitest unit & OpenXML integration tests
npm test

# Build ESM (dist/index.mjs + dist/index.d.ts) and IIFE browser bundle (dist/extended-pptxjs.bundle.js)
npm run build
```

---

## License

Licensed under the **GNU General Public License v3.0 (GPL-3.0)**. See [LICENSE](./LICENSE) for details.