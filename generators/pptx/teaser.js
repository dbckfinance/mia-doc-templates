import PptxGenJS from 'pptxgenjs';
import fs from 'node:fs';
import { BRAND } from '../../shared/branding.js';

const inputPath = process.argv[2] || './input.json';
const input = JSON.parse(fs.readFileSync(inputPath, 'utf8'));

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'M&IA';

const slide = pptx.addSlide();
slide.background = { color: BRAND.primary };

// Header
slide.addText(input.headline || 'Investment Opportunity', {
  x: 0.5, y: 0.3, w: '90%', h: 0.8,
  fontSize: 22, color: 'FFFFFF', bold: true,
});

// Metrics box
const metrics = input.metrics || { revenue: '€XX m', ebitda: '€XX m', growth: 'XX%', margin: 'XX%' };
slide.addShape(pptx.ShapeType.rect, {
  x: 0.5, y: 1.3, w: '90%', h: 1.2,
  fill: { color: BRAND.accent.replace('#', '') }, rectRadius: 0.05,
});
const metricsText = `Revenue: ${metrics.revenue}  |  EBITDA: ${metrics.ebitda}  |  Growth: ${metrics.growth}  |  Margin: ${metrics.margin}`;
slide.addText(metricsText, {
  x: 0.7, y: 1.5, w: '85%', h: 0.8,
  fontSize: 16, color: 'FFFFFF', align: 'center', valign: 'middle',
});

// Description
slide.addText(input.description || '[Business description]', {
  x: 0.5, y: 2.8, w: '90%', h: 1.5,
  fontSize: 13, color: 'DDDDDD', valign: 'top',
});

// Highlights
const highlights = input.highlights || ['[Highlight 1]', '[Highlight 2]', '[Highlight 3]'];
highlights.forEach((hl, i) => {
  slide.addText(`• ${hl}`, {
    x: 0.7, y: 4.5 + i * 0.45, w: '85%', h: 0.4,
    fontSize: 12, color: 'CCCCCC',
  });
});

const outputName = input.fileName || 'teaser';
await pptx.writeFile({ fileName: `./${outputName}.pptx` });
console.log(`OK: wrote ${outputName}.pptx`);
