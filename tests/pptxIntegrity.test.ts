import path from 'node:path';
import JSZip from 'jszip';
import { DOMParser, type Element as XmlElement } from '@xmldom/xmldom';
import { describe, expect, it } from 'vitest';
import { ExtendedPptxJS } from '../src/core/ExtendedPptxJS';
import { parseCssToShapeOptions } from '../src/parsers/cssParser';

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

function resolveRelsTarget(relsPath: string, target: string): string {
  if (target.startsWith('/')) {
    return target.slice(1);
  }
  // e.g. ppt/slides/_rels/slide1.xml.rels -> base folder is ppt/slides
  const relsDir = path.posix.dirname(relsPath);
  const ownerDir = relsDir.endsWith('/_rels')
    ? relsDir.slice(0, -'/_rels'.length)
    : relsDir === '_rels'
      ? ''
      : relsDir;
  return path.posix.normalize(ownerDir ? path.posix.join(ownerDir, target) : target);
}

describe('PPTX OpenXML structural integrity', () => {
  it('generates a multi-slide PPTX with valid XML, resolved .rels targets, and strict spPr child ordering', async () => {
    const ext = new ExtendedPptxJS();

    // Slide 1: Aurora glow, multi-stop gradient, exact pill radii, native translucent glass card
    const slide1 = ext.addSlide();

    ext.addExtendedShape(slide1, {
      xPx: 0,
      yPx: 0,
      wPx: 1920,
      hPx: 1080,
      shapeType: 'rect',
      gradient: {
        angleDeg: 135,
        stops: [
          { color: '070B14', position: 0 },
          { color: '0D1324', position: 55 },
          { color: '111827', position: 100 },
        ],
      },
    });

    // Aurora soft-edge glow blob
    ext.addExtendedShape(slide1, {
      xPx: 1200,
      yPx: -120,
      wPx: 680,
      hPx: 680,
      shapeType: 'ellipse',
      fillColor: '8B5CF6',
      fillOpacity: 0.42,
      softEdgePt: 45,
    });

    // Exact pill badge
    ext.addExtendedShape(slide1, {
      xPx: 84,
      yPx: 60,
      wPx: 300,
      hPx: 44,
      cornerRadiusPx: 999,
      gradient: {
        angleDeg: 120,
        stops: [
          { color: 'F472B6', position: 0 },
          { color: 'FDE047', position: 100 },
        ],
      },
    });

    // Native translucent glass card
    const glassOpts = parseCssToShapeOptions(
      {
        backgroundColor: 'rgba(15, 23, 42, 0.66)',
        borderRadius: '28px',
        border: '1.5px solid rgba(148, 163, 184, 0.28)',
        boxShadow: '0 24px 48px rgba(0, 0, 0, 0.45)',
      },
      { xPx: 84, yPx: 260, wPx: 540, hPx: 420 }
    );
    ext.addExtendedShape(slide1, glassOpts);

    // Slide 2: Combo gradient line/area chart, gradient column chart, donut chart, and table
    const slide2 = ext.addSlide();

    ext.addGradientLineChart(slide2, {
      xPx: 84,
      yPx: 180,
      wPx: 540,
      hPx: 360,
      seriesName: 'Throughput',
      labels: ['JAN', 'MAR', 'MAY', 'JUL', 'SEP', 'DEC'],
      values: [28, 54, 46, 82, 74, 98],
      lineColor: '#34D399',
      lineWidthPt: 3.5,
      showMarkers: true,
      markerSize: 7,
    });

    ext.addGradientBarChart(slide2, {
      xPx: 680,
      yPx: 180,
      wPx: 540,
      hPx: 360,
      seriesName: 'Revenue',
      labels: ['Q1', 'Q2', 'Q3', 'Q4'],
      values: [42, 68, 81, 96],
      barColor: '#6366F1',
      showDataLabels: true,
    });

    ext.addDonutChart(slide2, {
      xPx: 1280,
      yPx: 180,
      wPx: 360,
      hPx: 360,
      seriesName: 'Capacity',
      labels: ['Utilized', 'Available'],
      values: [78.4, 21.6],
      colors: ['38BDF8', 'E2E8F0'],
      holeSize: 70,
    });

    slide2.addTable(
      [
        [
          { text: 'Metric', options: { bold: true, fill: { color: 'EEF2FF' } } },
          { text: 'Value', options: { bold: true, align: 'right', fill: { color: 'EEF2FF' } } },
        ],
        [
          { text: 'Uptime' },
          { text: '99.98%', options: { align: 'right' } },
        ],
      ],
      {
        x: ext.pxToIn(84),
        y: ext.pxToIn(600),
        w: ext.pxToIn(1752),
        colW: [ext.pxToIn(1200), ext.pxToIn(552)],
      }
    );

    const arrayBuffer = await ext.exportBuffer();
    const zip = await JSZip.loadAsync(Buffer.from(arrayBuffer));

    const xmlErrors: string[] = [];
    const entryNames = new Set(Object.keys(zip.files).filter((n) => !zip.files[n].dir));

    for (const entryName of entryNames) {
      if (!entryName.endsWith('.xml') && !entryName.endsWith('.rels')) continue;

      const xmlText = await zip.file(entryName)!.async('string');
      const parser = new DOMParser({
        onError: (level, msg) => {
          if (level === 'error' || level === 'fatalError') {
            xmlErrors.push(`${entryName} [${level}]: ${msg}`);
          }
        },
      });
      const doc = parser.parseFromString(xmlText, 'application/xml');
      expect(doc.documentElement, `Missing root element in ${entryName}`).toBeTruthy();

      // Verify .rels Targets resolve to real ZIP entries
      if (entryName.endsWith('.rels')) {
        const rels = doc.getElementsByTagName('Relationship');
        for (let i = 0; i < rels.length; i++) {
          const rel = rels.item(i)!;
          if (rel.getAttribute('TargetMode') === 'External') continue;
          const target = rel.getAttribute('Target');
          expect(target).toBeTruthy();
          const resolved = resolveRelsTarget(entryName, target!);
          expect(
            entryNames.has(resolved),
            `Broken .rels target in ${entryName}: "${target}" -> "${resolved}"`
          ).toBe(true);
        }
      }

      // Verify strict OOXML child sequence on every <p:spPr> and <c:spPr>
      const allElements = doc.getElementsByTagName('*');
      for (let i = 0; i < allElements.length; i++) {
        const el = allElements.item(i)!;
        if (getLocalName(el) !== 'spPr') continue;

        const childOrderIndices: number[] = [];
        const childNames: string[] = [];
        for (let j = 0; j < el.childNodes.length; j++) {
          const child = el.childNodes.item(j);
          if (child && child.nodeType === 1) {
            const name = getLocalName(child as XmlElement);
            childNames.push(name);
            const idx = SP_PR_ORDER.indexOf(name);
            expect(idx, `Unexpected child <${name}> in ${entryName}`).toBeGreaterThanOrEqual(0);
            childOrderIndices.push(idx);
          }
        }

        const sorted = [...childOrderIndices].sort((a, b) => a - b);
        expect(
          childOrderIndices,
          `Invalid spPr child sequence in ${entryName}: [${childNames.join(', ')}]`
        ).toEqual(sorted);
      }
    }

    expect(xmlErrors).toEqual([]);

    // Verify combo chart series indices and gradient patching in chart1.xml
    const chart1Xml = await zip.file('ppt/charts/chart1.xml')!.async('string');
    expect(chart1Xml).toContain('<c:idx val="0"/>');
    expect(chart1Xml).toContain('<c:idx val="1"/>');
    expect(chart1Xml).toContain('Sheet1!$B$1');
    expect(chart1Xml).toContain('Sheet1!$C$1');
    expect(chart1Xml).toContain('<a:gradFill rotWithShape="1">');
    expect(chart1Xml).not.toContain('2094734556');

    // Verify column chart gradient fill in chart2.xml
    const chart2Xml = await zip.file('ppt/charts/chart2.xml')!.async('string');
    expect(chart2Xml).toContain('<c:barChart>');
    expect(chart2Xml).toContain('<a:gradFill rotWithShape="1">');

    // Verify donut chart holeSize in chart3.xml
    const chart3Xml = await zip.file('ppt/charts/chart3.xml')!.async('string');
    expect(chart3Xml).toContain('<c:doughnutChart>');
    expect(chart3Xml).toContain('<c:holeSize val="70"/>');
  });
});
