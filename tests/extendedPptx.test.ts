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
    expect(shapes[1]).toContain('<a:softEdge rad="571500"/>');
    expect(shapes[2]).toContain('<a:custGeom>');
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
