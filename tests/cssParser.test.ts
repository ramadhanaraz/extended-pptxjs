import { describe, expect, it } from 'vitest';
import {
  parseBorder,
  parseBorderRadius,
  parseBoxShadow,
  parseCssColor,
  parseCssToShapeOptions,
  parseFilterBlur,
  parseLinearGradient,
  resolveFontFamily,
} from '../src/parsers/cssParser';

describe('parseCssColor', () => {
  it('parses 3-char, 4-char, 6-char, and 8-char hex colors', () => {
    expect(parseCssColor('#abc')).toEqual({ hex: 'AABBCC', opacity: 1 });
    expect(parseCssColor('#f088')).toEqual({
      hex: 'FF0088',
      opacity: Number((0x88 / 255).toFixed(4)),
    });
    expect(parseCssColor('#3b82f6')).toEqual({ hex: '3B82F6', opacity: 1 });
    expect(parseCssColor('#3b82f680')).toEqual({
      hex: '3B82F6',
      opacity: Number((128 / 255).toFixed(4)),
    });
  });

  it('parses rgb() and rgba() in both comma and space-slash syntax', () => {
    expect(parseCssColor('rgb(15, 23, 42)')).toEqual({ hex: '0F172A', opacity: 1 });
    expect(parseCssColor('rgba(15, 23, 42, 0.78)')).toEqual({
      hex: '0F172A',
      opacity: 0.78,
    });
    expect(parseCssColor('rgb(255 0 128 / 50%)')).toEqual({
      hex: 'FF0080',
      opacity: 0.5,
    });
  });

  it('parses hsl(), hsla(), transparent, and named colors', () => {
    expect(parseCssColor('hsl(0, 100%, 50%)')).toEqual({ hex: 'FF0000', opacity: 1 });
    expect(parseCssColor('hsla(240, 100%, 50%, 0.4)')).toEqual({
      hex: '0000FF',
      opacity: 0.4,
    });
    expect(parseCssColor('transparent')).toEqual({ hex: '000000', opacity: 0 });
    expect(parseCssColor('white')).toEqual({ hex: 'FFFFFF', opacity: 1 });
    expect(parseCssColor('none')).toBeUndefined();
  });
});

describe('parseLinearGradient', () => {
  it('parses multi-stop linear-gradient with angle and rgba opacity', () => {
    const grad = parseLinearGradient(
      'linear-gradient(135deg, #f472b6 0%, rgba(251, 113, 133, 0.65) 52%, #fde047 100%)'
    );
    expect(grad).toEqual({
      angleDeg: 135,
      stops: [
        { color: 'F472B6', position: 0 },
        { color: 'FB7185', position: 52, opacity: 0.65 },
        { color: 'FDE047', position: 100 },
      ],
    });
  });

  it('parses directional keywords and interpolates missing stop positions', () => {
    const grad = parseLinearGradient('linear-gradient(to right, #6366f1, #a855f7, #ec4899)');
    expect(grad).toEqual({
      angleDeg: 90,
      stops: [
        { color: '6366F1', position: 0 },
        { color: 'A855F7', position: 50 },
        { color: 'EC4899', position: 100 },
      ],
    });
  });

  it('defaults omitted angle to 180deg (to bottom) and supports turn units', () => {
    const defaultDir = parseLinearGradient('linear-gradient(#070b14, #111827)');
    expect(defaultDir?.angleDeg).toBe(180);
    expect(defaultDir?.stops).toEqual([
      { color: '070B14', position: 0 },
      { color: '111827', position: 100 },
    ]);

    const turnDir = parseLinearGradient('linear-gradient(0.25turn, #000000 0%, #ffffff 100%)');
    expect(turnDir?.angleDeg).toBe(90);
  });
});

describe('parseBorderRadius', () => {
  it('parses px values, multi-corner averages, and percentage radii', () => {
    expect(parseBorderRadius('20px')).toBe(20);
    expect(parseBorderRadius('28px 12px 24px 12px')).toBe(19);
    expect(parseBorderRadius('50%', 120, 120)).toBe(60);
    expect(parseBorderRadius('0px')).toBe(0);
    expect(parseBorderRadius('none')).toBeUndefined();
  });
});

describe('parseBoxShadow', () => {
  it('parses inline box-shadow and converts blur px to points (0.5pt/px)', () => {
    expect(parseBoxShadow('0 18px 40px -10px rgba(0, 0, 0, 0.75)')).toEqual({
      color: '000000',
      opacity: 0.75,
      blurPt: 20,
      offsetXPx: 0,
      offsetYPx: 18,
    });
  });

  it('parses browser getComputedStyle color-first box-shadow order', () => {
    expect(parseBoxShadow('rgba(99, 102, 241, 0.25) 4px 10px 24px 0px')).toEqual({
      color: '6366F1',
      opacity: 0.25,
      blurPt: 12,
      offsetXPx: 4,
      offsetYPx: 10,
    });
  });

  it('handles zero-blur neo-brutalist shadows and skips inset shadows', () => {
    expect(parseBoxShadow('5px 5px 0px #0f172a')).toEqual({
      color: '0F172A',
      opacity: 1,
      blurPt: 0,
      offsetXPx: 5,
      offsetYPx: 5,
    });

    expect(
      parseBoxShadow('inset 0 2px 4px rgba(255,255,255,0.2), 0 8px 20px rgba(15,23,42,0.3)')
    ).toEqual({
      color: '0F172A',
      opacity: 0.3,
      blurPt: 10,
      offsetXPx: 0,
      offsetYPx: 8,
    });
  });
});

describe('parseBorder & parseFilterBlur', () => {
  it('parses border width, style, and color with opacity', () => {
    expect(parseBorder('1.5px solid rgba(129, 140, 248, 0.32)')).toEqual({
      color: '818CF8',
      widthPt: 0.75,
      opacity: 0.32,
    });
    expect(parseBorder('0px none rgb(0, 0, 0)')).toBeUndefined();
  });

  it('parses CSS filter blur into softEdgePt', () => {
    expect(parseFilterBlur('blur(48px) brightness(1.1)')).toBe(24);
    expect(parseFilterBlur('none')).toBeUndefined();
  });
});

describe('parseCssToShapeOptions', () => {
  it('combines gradient, border-radius, shadow, border, and element opacity into ExtendedShapeOptions', () => {
    const opts = parseCssToShapeOptions(
      {
        backgroundImage: 'linear-gradient(135deg, #3b82f6 0%, #8b5cf6 100%)',
        borderRadius: '24px',
        boxShadow: '0 20px 40px rgba(0, 0, 0, 0.5)',
        border: '2px solid rgba(255, 255, 255, 0.4)',
        opacity: '0.8',
      },
      { xPx: 100, yPx: 200, wPx: 600, hPx: 320 }
    );

    expect(opts).toEqual({
      xPx: 100,
      yPx: 200,
      wPx: 600,
      hPx: 320,
      shapeType: 'roundRect',
      cornerRadiusPx: 24,
      gradient: {
        angleDeg: 135,
        stops: [
          { color: '3B82F6', position: 0, opacity: 0.8 },
          { color: '8B5CF6', position: 100, opacity: 0.8 },
        ],
      },
      shadow: {
        color: '000000',
        opacity: 0.5,
        blurPt: 20,
        offsetXPx: 0,
        offsetYPx: 20,
      },
      border: {
        color: 'FFFFFF',
        widthPt: 1,
        opacity: 0.32,
      },
    });
  });

  it('maps 50% border-radius circles to ellipse and solid rgba fills with blur to softEdgePt', () => {
    const opts = parseCssToShapeOptions(
      {
        backgroundColor: 'rgba(56, 189, 248, 0.35)',
        borderRadius: '50%',
        filter: 'blur(60px)',
      },
      { xPx: 50, yPx: 50, wPx: 400, hPx: 400 }
    );

    expect(opts).toEqual({
      xPx: 50,
      yPx: 50,
      wPx: 400,
      hPx: 400,
      shapeType: 'ellipse',
      fillColor: '38BDF8',
      fillOpacity: 0.35,
      softEdgePt: 30,
    });
  });
});

describe('resolveFontFamily', () => {
  it('extracts clean primary font family and strips quotes while ignoring generic fallbacks', () => {
    expect(
      resolveFontFamily('"Plus Jakarta Sans", -apple-system, BlinkMacSystemFont, sans-serif')
    ).toBe('Plus Jakarta Sans');
    expect(resolveFontFamily("'Space Grotesk', serif")).toBe('Space Grotesk');
    expect(resolveFontFamily('-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif')).toBe(
      'Segoe UI'
    );
  });

  it('maps monospace stacks to Consolas by default and supports custom fontMap overrides', () => {
    expect(
      resolveFontFamily('ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace')
    ).toBe('Consolas');
    expect(resolveFontFamily('monospace')).toBe('Consolas');
    expect(resolveFontFamily('sans-serif')).toBe('Arial');
    expect(resolveFontFamily('serif')).toBe('Georgia');

    expect(
      resolveFontFamily('"Plus Jakarta Sans", sans-serif', {
        'Plus Jakarta Sans': 'Calibri',
      })
    ).toBe('Calibri');
  });
});
