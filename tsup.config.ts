import { defineConfig } from 'tsup';

export default defineConfig([
  // 1. ESM module + TypeScript declarations
  {
    entry: ['src/index.ts'],
    format: ['esm'],
    dts: true,
    clean: true,
    sourcemap: true,
  },
  // 2. Self-contained single-file browser bundle (pptxgenjs + jszip inlined)
  {
    entry: { 'extended-pptxjs.bundle': 'src/browser.ts' },
    format: ['iife'],
    globalName: 'ExtendedPptxJSLib',
    noExternal: ['pptxgenjs', 'jszip'],
    minify: true,
    platform: 'browser',
  },
]);