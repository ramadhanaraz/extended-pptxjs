import { defineConfig } from 'tsup';

export default defineConfig([
  // 1. ESM module + TypeScript declarations (dist/index.mjs + dist/index.d.ts)
  {
    entry: ['src/index.ts'],
    format: ['esm'],
    dts: true,
    clean: true,
    sourcemap: true,
    outExtension: () => ({ js: '.mjs' }),
  },
  // 2. Self-contained single-file browser bundle (dist/extended-pptxjs.bundle.js)
  {
    entry: { 'extended-pptxjs.bundle': 'src/browser.ts' },
    format: ['iife'],
    globalName: 'ExtendedPptxJSLib',
    noExternal: ['pptxgenjs', 'jszip', '@xmldom/xmldom'],
    minify: true,
    platform: 'browser',
    outExtension: () => ({ js: '.js' }),
  },
]);