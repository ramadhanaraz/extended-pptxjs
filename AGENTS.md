# AGENTS.md — ExtendedPptxJS Library

## Project Overview
`extended-pptxjs` is a TypeScript wrapper and OpenXML post-processor built on `pptxgenjs` and `jszip`. It converts modern HTML/CSS presentation slides (`1920x1080` base coordinate system) into native, editable `.pptx` decks (`13.333" x 7.5"` 16:9 widescreen) for use in Digitan's Guide and executive slide decks.

## Core Architectural Rules
1. **Editable Vectors First:** Never rasterize text or standard geometric shapes. Tag `pptxgenjs` shapes via `objectName` and post-process the generated `.pptx` zip archive using `JSZip` to inject native OOXML `<a:gradFill>` multi-stop gradients, alpha transparency channels, custom corner radii, `<a:outerShdw>` drop shadows, and `<a:softEdge>` blurs inside `<p:spPr>`.
2. **Hybrid Raster Fallback Only When Necessary:** Only rasterize CSS/SVG features that have no OpenXML equivalent—specifically `backdrop-filter: blur()` frosted glass cards, multi-layer mesh blurs, and complex multi-filter SVGs. Render those onto an offscreen HTML5 Canvas at 3x DPI (`5760x3240` equivalent) and embed them as transparent PNGs.
3. **Coordinate Normalization:** All public layout APIs accept `1920x1080` pixel coordinates (`xPx`, `yPx`, `wPx`, `hPx`) and normalize them to slide inches (`13.333" x 7.5"`).
4. **Zero-Dependency Browser Bundle:** Every `npm run build` must succeed for both `dist/index.mjs` (ESM + `.d.ts`) and `dist/extended-pptxjs.bundle.js` (IIFE bundle with `pptxgenjs` and `jszip` inlined via `noExternal`).
5. **Offline & Firewall Safety:** Keep zero external network calls at runtime so the bundle works offline in air-gapped or firewalled environments.
6. **Verification Gate:** Run `npm run typecheck`, `npm test`, and `npm run build` before finishing any task.

## Directory Structure
- `src/core/ExtendedPptxJS.ts` — Core wrapper, px-to-inch math, and `JSZip` OpenXML patcher.
- `src/parsers/cssParser.ts` — Parses CSS `linear-gradient()`, `rgba()`, border-radius, and `box-shadow` into OOXML-compatible structs.
- `src/renderers/hybridAssetRenderer.ts` — Converts geometric SVGs to `<a:custGeom>` and frosted-glass/complex layers to 3x DPI transparent PNGs.
- `src/walker/HtmlSlideConverter.ts` — Traverses a `1920x1080` HTML slide container in paint/z-index order and compiles slides.
- `src/index.ts` — ESM library entry point.
- `src/browser.ts` — Browser IIFE entry point exposing `window.ExtendedPptxJS`.
- `playground/index.html` — Visual test bench with dark/light Aurora gradients, frosted glass cards, and SVG icons.