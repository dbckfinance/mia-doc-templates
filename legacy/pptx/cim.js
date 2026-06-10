import PptxGenJS from 'pptxgenjs';
import fs from 'node:fs';
import { BRAND } from '../../shared/branding.js';

const inputPath = process.argv[2] || './input.json';
const input = JSON.parse(fs.readFileSync(inputPath, 'utf8'));

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'M&IA';

// Cover slide
const cover = pptx.addSlide();
cover.background = { color: BRAND.primary };
cover.addText(input.companyName || 'Confidential Information Memorandum', {
  x: 0.5, y: 2.0, w: '90%', h: 1.5,
  fontSize: 32, color: 'FFFFFF', bold: true, align: 'center',
});
cover.addText(BRAND.confidentiality, {
  x: 0.5, y: 4.5, w: '90%', h: 0.5,
  fontSize: 12, color: 'AAAAAA', align: 'center',
});

// Executive Summary
const exec = pptx.addSlide();
exec.addText('Executive Summary', {
  x: 0.5, y: 0.3, w: '90%', h: 0.6,
  fontSize: 24, bold: true, color: BRAND.primary.replace('#', ''),
});
exec.addText(input.executiveSummary || '[Executive summary content]', {
  x: 0.5, y: 1.2, w: '90%', h: 4.0,
  fontSize: 14, color: '333333', valign: 'top',
});

// Investment Highlights
const highlights = pptx.addSlide();
highlights.addText('Investment Highlights', {
  x: 0.5, y: 0.3, w: '90%', h: 0.6,
  fontSize: 24, bold: true, color: BRAND.primary.replace('#', ''),
});
const hlItems = input.investmentHighlights || ['[Highlight 1]', '[Highlight 2]', '[Highlight 3]'];
hlItems.forEach((hl, i) => {
  highlights.addText(`${i + 1}. ${hl}`, {
    x: 0.8, y: 1.2 + i * 0.8, w: '85%', h: 0.7,
    fontSize: 14, color: '333333',
  });
});

const outputName = input.fileName || 'cim';
await pptx.writeFile({ fileName: `./${outputName}.pptx` });
console.log(`OK: wrote ${outputName}.pptx`);
