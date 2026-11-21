import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { JSDOM } from 'jsdom';
import JSZip from 'jszip';
import { DOMParser, type Element as XmlElement } from '@xmldom/xmldom';
import { describe, expect, it } from 'vitest';
import { ExtendedPptxJS } from '../src/core/ExtendedPptxJS';
import { HtmlSlideConverter } from '../src/walker/HtmlSlideConverter';

const SP_PR_ORDER = [
  'xfrm',
  'custGeom',
  'prstGeom',
  'noFill',
  'solidFill',
  'gradFill',
  'blipFill',
  'pattFill',
  'grpFill',
  'ln',
  'effectLst',
  'effectDag',
  'scene3d',
  'sp3d',
  'extLst',
];

function getLocalName(el: XmlElement): string {
  if (el.localName) return el.localName;
  const name = el.tagName ?? el.nodeName ?? '';
  const idx = name.indexOf(':');
  return idx >= 0 ? name.slice(idx + 1) : name;
}

function makeDomRect(x: number, y: number, w: number, h: number): DOMRect {
  return {
    x,
    y,
    width: w,
    height: h,
    top: y,
    left: x,
    right: x + w,
    bottom: y + h,
    toJSON: () => ({}),
  } as DOMRect;
}

describe('HtmlSlideConverter multi-slide DOM integration', () => {
  it('handles scaled slides, hidden display:none slides, single-line +14% expansion, rich inline text, single-sided borders, background-clip:text, ::before, charts, and tables', async () => {
    const dom = new JSDOM(
      `<!DOCTYPE html><html><head><style>
        .slide { width: 1920px; height: 1080px; }
        .hidden-slide { display: none; }
        .accent-card { position: relative; background: #0f172a; border-radius: 20px; }
        .accent-card::before {
          content: "";
          position: absolute;
          top: 0px;
          left: 0px;
          right: 0px;
          height: 4px;
          background: linear-gradient(90deg, #38bdf8, #818cf8);
        }
        .grad-text {
          background-image: linear-gradient(90deg, #38bdf8 0%, #f472b6 100%);
          -webkit-background-clip: text;
          background-clip: text;
          -webkit-text-fill-color: transparent;
        }
      </style></head><body>
        <section id="slide-1" class="slide" style="transform: scale(0.75); background: #070b14;">
          <h1 id="headline" style="font-family: 'Plus Jakarta Sans', sans-serif; font-size: 48px; line-height: 56px; color: #f8fafc; text-align: left;">
            <span id="clip-span" class="grad-text">Aurora</span> Executive Deck
          </h1>
          <div id="rich-card" class="accent-card">
            <p id="rich-para" style="font-size: 22px; line-height: 30px; color: #cbd5e1;">
              Built with <strong>native vectors</strong> and <span style="color: #34d399;">inline color</span>.<br/>Second line.
            </p>
          </div>
          <footer id="footer" style="border-top: 2px solid #334155;">
            <span id="footer-label" style="font-family: ui-monospace, monospace; font-size: 16px; color: #94a3b8;">SLIDE 01</span>
          </footer>
        </section>

        <section id="slide-2" class="slide hidden-slide" style="background: #f8fafc;">
          <div id="chart-box" data-pptx-chart='{"type":"gradient-line","seriesName":"Growth","labels":["Q1","Q2","Q3"],"values":[20,55,90],"lineColor":"#0EA5E9"}'></div>
          <table id="data-table">
            <thead><tr><th style="text-align:left;">Metric</th><th style="text-align:right;">Score</th></tr></thead>
            <tbody><tr><td>Reliability</td><td style="text-align:right;">99.9%</td></tr></tbody>
          </table>
        </section>
      </body></html>`,
      { pretendToBeVisual: true }
    );

    const win = dom.window;
    const doc = win.document;

    // Assign mock offsetWidth/offsetHeight & getBoundingClientRect respecting scale(0.75) and display:none
    const slide1 = doc.getElementById('slide-1')!;
    const slide2 = doc.getElementById('slide-2')!;
    const headline = doc.getElementById('headline')!;
    const richCard = doc.getElementById('rich-card')!;
    const richPara = doc.getElementById('rich-para')!;
    const footer = doc.getElementById('footer')!;
    const footerLabel = doc.getElementById('footer-label')!;
    const chartBox = doc.getElementById('chart-box')!;
    const dataTable = doc.getElementById('data-table') as HTMLTableElement;

    const layoutRects = new Map<Element, [number, number, number, number]>([
      [slide1, [0, 0, 1920, 1080]],
      [headline, [80, 60, 600, 56]],
      [richCard, [80, 180, 800, 300]],
      [richPara, [110, 220, 700, 64]],
      [footer, [80, 960, 1760, 48]],
      [footerLabel, [80, 976, 120, 20]],
      [slide2, [0, 0, 1920, 1080]],
      [chartBox, [80, 120, 800, 360]],
      [dataTable, [80, 540, 800, 120]],
    ]);

    for (const tr of Array.from(dataTable.rows)) {
      layoutRects.set(tr, [80, 540, 800, 60]);
      for (const cell of Array.from(tr.cells)) {
        layoutRects.set(cell, [80, 540, 400, 60]);
      }
    }

    const isHidden = (el: Element): boolean => {
      let cur: Element | null = el;
      while (cur) {
        if (win.getComputedStyle(cur).display === 'none') return true;
        cur = cur.parentElement;
      }
      return false;
    };

    for (const [el, [x, y, w, h]] of layoutRects.entries()) {
      Object.defineProperty(el, 'offsetWidth', { configurable: true, get: () => w });
      Object.defineProperty(el, 'offsetHeight', { configurable: true, get: () => h });
      (el as HTMLElement).getBoundingClientRect = () => {
        if (isHidden(el)) return makeDomRect(0, 0, 0, 0);
        const scale = slide1.contains(el) ? 0.75 : 1;
        return makeDomRect(x * scale, y * scale, w * scale, h * scale);
      };
    }

    // Mock Range.prototype.getBoundingClientRect and getClientRects
    win.Range.prototype.getBoundingClientRect = function () {
      const container =
        this.startContainer.nodeType === 3
          ? this.startContainer.parentElement
          : (this.startContainer as Element);
      const target = container?.closest('[id]') ?? container;
      if (target && layoutRects.has(target)) {
        const [x, y, w, h] = layoutRects.get(target)!;
        const scale = slide1.contains(target) ? 0.75 : 1;
        return makeDomRect(x * scale, y * scale, w * scale, h * scale);
      }
      return makeDomRect(0, 0, 0, 0);
    };

    win.Range.prototype.getClientRects = function () {
      const rect = this.getBoundingClientRect();
      return [rect] as unknown as DOMRectList;
    };

    // Bind globals temporarily for HtmlSlideConverter
    const origWin = globalThis.window;
    const origDoc = globalThis.document;
    const origHtmlEl = globalThis.HTMLElement;
    const origSvgEl = globalThis.SVGSVGElement;
    const origTableEl = globalThis.HTMLTableElement;
    const origImgEl = globalThis.HTMLImageElement;
    const origNode = globalThis.Node;

    Object.assign(globalThis, {
      window: win,
      document: doc,
      HTMLElement: win.HTMLElement,
      SVGSVGElement: win.SVGSVGElement,
      HTMLTableElement: win.HTMLTableElement,
      HTMLImageElement: win.HTMLImageElement,
      Node: win.Node,
    });

    try {
      const ext = new ExtendedPptxJS();
      const converter = new HtmlSlideConverter(ext, { glassMode: 'native-translucent' });
      await converter.convertDeck([slide1, slide2]);

      // Hidden slide2 must have its original display:none restored
      expect(win.getComputedStyle(slide2).display).toBe('none');

      const buffer = await ext.exportBuffer();
      const zip = await JSZip.loadAsync(Buffer.from(buffer));

      // 1. Validate XML syntax and spPr strict child sequence across entire archive
      const entries = new Set(Object.keys(zip.files).filter((n) => !zip.files[n].dir));
      const xmlErrors: string[] = [];
      for (const name of entries) {
        if (!name.endsWith('.xml') && !name.endsWith('.rels')) continue;
        const xml = await zip.file(name)!.async('string');
        const parsed = new DOMParser({
          onError: (lvl, msg) => {
            if (lvl === 'error' || lvl === 'fatalError') xmlErrors.push(`${name}: ${msg}`);
          },
        }).parseFromString(xml, 'application/xml');

        const all = parsed.getElementsByTagName('*');
        for (let i = 0; i < all.length; i++) {
          const el = all.item(i)!;
          if (getLocalName(el) !== 'spPr') continue;
          const idxs: number[] = [];
          for (let j = 0; j < el.childNodes.length; j++) {
            const c = el.childNodes.item(j);
            if (c && c.nodeType === 1) idxs.push(SP_PR_ORDER.indexOf(getLocalName(c as XmlElement)));
          }
          expect(idxs).toEqual([...idxs].sort((a, b) => a - b));
        }
      }
      expect(xmlErrors).toEqual([]);

      const s1Xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
      const s2Xml = await zip.file('ppt/slides/slide2.xml')!.async('string');

      // Single-line headline: margin:0, wrap="none", <a:normAutofit/>, +14% width (600 * 1.14 = 684px -> 684/1920*13.333 in)
      expect(s1Xml).toContain('wrap="none"');
      expect(s1Xml).toContain('lIns="0" tIns="0" rIns="0" bIns="0"');
      expect(s1Xml).toContain('<a:normAutofit/>');
      const expectedExpandedCx = Math.round(ext.pxToIn(600 * 1.14) * 914400);
      expect(s1Xml).toContain(`cx="${expectedExpandedCx}"`);

      // background-clip:text extracted first stop #38BDF8 for "Aurora" run
      expect(s1Xml).toContain('val="38BDF8"');
      expect(s1Xml).toContain('<a:t>Aurora</a:t>');

      // Rich inline text runs preserved (bold "native vectors" and green "#34D399" "inline color")
      expect(s1Xml).toContain('<a:t>native vectors</a:t>');
      expect(s1Xml).toContain('val="34D399"');
      expect(s1Xml).toContain('<a:t>inline color</a:t>');

      // Single-sided border-top footer emitted as prst="line", not a 4-sided rect border
      expect(s1Xml).toContain('<a:prstGeom prst="line">');

      // ::before accent bar emitted with #818CF8 gradient stop
      expect(s1Xml).toContain('val="818CF8"');

      // Slide 2 (previously display:none) converted with native chart + table + Consolas font
      expect(s1Xml).toContain('typeface="Consolas"');
      expect(s2Xml).toContain('<a:tbl>');
      expect(zip.file('ppt/charts/chart1.xml')).toBeTruthy();
    } finally {
      Object.assign(globalThis, {
        window: origWin,
        document: origDoc,
        HTMLElement: origHtmlEl,
        SVGSVGElement: origSvgEl,
        HTMLTableElement: origTableEl,
        HTMLImageElement: origImgEl,
        Node: origNode,
      });
    }
  });

  it('attaches constructible globals on window and globalThis without colliding with ExtendedPptxJSLib', () => {
    const bundlePath = path.resolve(__dirname, '../dist/extended-pptxjs.bundle.js');
    const code = fs.readFileSync(bundlePath, 'utf8');
    const sandbox: Record<string, unknown> = {
      console,
      Blob,
      URL,
      Uint8Array,
      ArrayBuffer,
      TextEncoder,
      TextDecoder,
      setTimeout,
      clearTimeout,
    };
    sandbox.window = sandbox;
    sandbox.globalThis = sandbox;
    sandbox.self = sandbox;

    vm.runInNewContext(code, sandbox);

    expect(typeof sandbox.ExtendedPptxJS).toBe('function');
    expect(typeof sandbox.HtmlSlideConverter).toBe('function');
    expect(typeof sandbox.cssParser).toBe('object');
    expect(typeof sandbox.ExtendedPptxJSLib).toBe('object');

    const ExtCtor = sandbox.ExtendedPptxJS as new () => ExtendedPptxJS;
    const ConvCtor = sandbox.HtmlSlideConverter as new (ext: ExtendedPptxJS) => HtmlSlideConverter;
    const instance = new ExtCtor();
    const conv = new ConvCtor(instance);
    expect(instance.baseWidthPx).toBe(1920);
    expect(conv).toBeDefined();
  });
});
