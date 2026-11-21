import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { ExtendedPptxJS } from '../src/core/ExtendedPptxJS';
import { parseCssToShapeOptions } from '../src/parsers/cssParser';
import {
  hasBackdropBlur,
  isComplexGradientOrMesh,
  parseBackdropFilter,
  parseRadialGradient,
} from '../src/renderers/hybridAssetRenderer';

describe('ExtendedPptxJS OpenXML post-processing', () => {
  it('injects native gradFill, outerShdw, softEdge, and custGeom into exported PPTX XML', async () => {
    const ext = new ExtendedPptxJS();
    const slide = ext.addSlide();

    const gradCardOpts = parseCssToShapeOptions(
      {
        backgroundImage: 'linear-gradient(135deg, #312e81 0%, #6d28d9 52%, #db2777 100%)',
        borderRadius: '28px',
        boxShadow: '0 24px 48px rgba(0, 0, 0, 0.5)',
        border: '2px solid rgba(255, 255, 255, 0.25)',
      },
      { xPx: 100, yPx: 100, wPx: 500, hPx: 300 }
    );
    ext.addExtendedShape(slide, gradCardOpts);

    const blobOpts = parseCssToShapeOptions(
      {
        backgroundColor: 'rgba(139, 92, 246, 0.42)',
        borderRadius: '50%',
        filter: 'blur(90px)',
      },
      { xPx: 1200, yPx: -100, wPx: 680, hPx: 680 }
    );
    ext.addExtendedShape(slide, blobOpts);

    ext.addExtendedShape(slide, {
      xPx: 700,
      yPx: 200,
      wPx: 100,
      hPx: 100,
      fillColor: '38BDF8',
      customGeometryXml:
        '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="100" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="100" y="50"/></a:lnTo><a:lnTo><a:pt x="0" y="100"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>',
    });

    const buffer = await ext.exportBuffer();
    const zip = await JSZip.loadAsync(buffer);
    const slideXml = await zip.file('ppt/slides/slide1.xml')?.async('string');

    expect(slideXml).toBeDefined();
    expect(slideXml).toContain('<a:gradFill rotWithShape="1">');
    expect(slideXml).toContain('val="6D28D9"');
    expect(slideXml).toContain('<a:outerShdw');
    expect(slideXml).toContain('<a:softEdge rad="571500"/>'); // 45pt * 12700
    expect(slideXml).toContain('<a:custGeom>');

    // Ensure patches landed on their own respective shapes
    const shapes = slideXml!.match(/<p:sp>[\s\S]*?<\/p:sp>/g) ?? [];
    expect(shapes).toHaveLength(3);
    expect(shapes[0]).toContain('<a:gradFill');
    expect(shapes[0]).toContain('<a:outerShdw');
    // 28px radius on 500x300 card -> round((28/300) * 100000) = 9333
    expect(shapes[0]).toContain('<a:gd name="adj" fmla="val 9333"/>');
    expect(shapes[1]).toContain('<a:softEdge rad="571500"/>');
    expect(shapes[2]).toContain('<a:custGeom>');
  });

  it('patches exact roundRect adj for wide pill badges, embeds TTF fonts, and builds gradient combo charts', async () => {
    const fakeTtfBytes = new Uint8Array([0x00, 0x01, 0x00, 0x00, 0x44, 0x49, 0x47, 0x49]);
    const ext = new ExtendedPptxJS({
      embedFonts: [{ name: 'Plus Jakarta Sans', data: fakeTtfBytes }],
    });
    const slide = ext.addSlide();

    // Wide pill badge: 320x48 with border-radius: 999px -> clamped to 50000
    ext.addExtendedShape(slide, {
      xPx: 80,
      yPx: 60,
      wPx: 320,
      hPx: 48,
      cornerRadiusPx: 999,
      fillColor: '6366F1',
    });

    // Rounded card: 400x200 with border-radius: 20px -> (20/200)*100000 = 10000
    ext.addExtendedShape(slide, {
      xPx: 80,
      yPx: 140,
      wPx: 400,
      hPx: 200,
      cornerRadiusPx: 20,
      fillColor: '0F172A',
    });

    // Native Gradient Area + Line Combo Chart
    ext.addGradientLineChart(slide, {
      xPx: 540,
      yPx: 140,
      wPx: 600,
      hPx: 320,
      seriesName: 'Throughput',
      labels: ['Q1', 'Q2', 'Q3', 'Q4'],
      values: [24, 58, 46, 92],
      lineColor: '#34D399',
      lineWidthPt: 3.5,
      showMarkers: true,
      markerSize: 8,
    });

    const buffer = await ext.exportBuffer();
    const zip = await JSZip.loadAsync(buffer);

    const slideXml = await zip.file('ppt/slides/slide1.xml')?.async('string');
    const shapes = slideXml!.match(/<p:sp>[\s\S]*?<\/p:sp>/g) ?? [];
    expect(shapes).toHaveLength(2);
    expect(shapes[0]).toContain('<a:gd name="adj" fmla="val 50000"/>');
    expect(shapes[1]).toContain('<a:gd name="adj" fmla="val 10000"/>');

    // Verify chart1.xml contains areaChart + lineChart, hidden axes, and vertical gradFill
    const chartXml = await zip.file('ppt/charts/chart1.xml')?.async('string');
    expect(chartXml).toBeDefined();
    expect(chartXml).toContain('<c:areaChart>');
    expect(chartXml).toContain('<c:lineChart>');
    expect(chartXml).toContain('<c:delete val="1"/>');
    expect(chartXml).toContain('<a:gradFill rotWithShape="1">');
    expect(chartXml).toContain('<a:lin ang="5400000" scaled="1"/>');
    expect(chartXml).toContain('<a:srgbClr val="34D399">');

    // Verify embedded font files and XML metadata
    const fontData = await zip.file('ppt/fonts/font1.fntdata')?.async('uint8array');
    expect(fontData).toEqual(fakeTtfBytes);

    const presXml = await zip.file('ppt/presentation.xml')?.async('string');
    expect(presXml).toContain('<p:embeddedFontLst>');
    expect(presXml).toContain('<p:font typeface="Plus Jakarta Sans"/>');

    const relsXml = await zip.file('ppt/_rels/presentation.xml.rels')?.async('string');
    expect(relsXml).toContain('Target="fonts/font1.fntdata"');

    const contentTypesXml = await zip.file('[Content_Types].xml')?.async('string');
    expect(contentTypesXml).toContain('Extension="fntdata"');
  });
});

describe('hybridAssetRenderer CSS inspection helpers', () => {
  it('detects backdrop-filter blur and parses blur + saturate values', () => {
    const glassStyle = {
      backdropFilter: 'blur(28px) saturate(180%)',
    };
    expect(hasBackdropBlur(glassStyle)).toBe(true);
    expect(parseBackdropFilter(glassStyle)).toEqual({ blurPx: 28, saturatePct: 180 });
    expect(hasBackdropBlur({ backdropFilter: 'none' })).toBe(false);
  });

  it('detects radial/multi-layer gradients and parses radial gradient stops', () => {
    expect(
      isComplexGradientOrMesh({
        backgroundImage: 'radial-gradient(circle at center, rgba(99,102,241,0.2) 0%, #070b14 85%)',
      })
    ).toBe(true);
    expect(
      isComplexGradientOrMesh({
        backgroundImage: 'linear-gradient(135deg, #000 0%, #fff 100%)',
      })
    ).toBe(false);

    const radial = parseRadialGradient(
      'radial-gradient(circle at 25% 75%, rgba(56, 189, 248, 0.4) 0%, #0f172a 90%)'
    );
    expect(radial).toEqual({
      cxRatio: 0.25,
      cyRatio: 0.75,
      stops: [
        { color: { hex: '38BDF8', opacity: 0.4 }, pos: 0 },
        { color: { hex: '0F172A', opacity: 1 }, pos: 0.9 },
      ],
    });
  });
});
