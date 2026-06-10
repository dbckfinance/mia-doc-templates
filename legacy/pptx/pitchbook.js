import PptxGenJS from 'pptxgenjs';
import fs from 'node:fs';
import { BRAND } from '../../shared/branding.js';

const inputPath = process.argv[2] || './input.json';
const input = JSON.parse(fs.readFileSync(inputPath, 'utf8'));

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'M&IA';

// Cover
const cover = pptx.addSlide();
cover.background = { color: BRAND.primary };
cover.addText(input.title || 'Pitchbook', {
  x: 0.5, y: 2.0, w: '90%', h: 1.0,
  fontSize: 30, color: 'FFFFFF', bold: true, align: 'center',
});
cover.addText(input.subtitle || 'Confidential', {
  x: 0.5, y: 3.2, w: '90%', h: 0.6,
  fontSize: 16, color: 'AAAAAA', align: 'center',
});

// Agenda
const agenda = pptx.addSlide();
agenda.addText('Agenda', {
  x: 0.5, y: 0.3, w: '90%', h: 0.6,
  fontSize: 24, bold: true, color: BRAND.primary.replace('#', ''),
});
const sections = input.sections || ['Situation Overview', 'Strategic Alternatives', 'Valuation', 'Process & Timeline'];
sections.forEach((s, i) => {
  agenda.addText(`${i + 1}. ${s}`, {
    x: 0.8, y: 1.2 + i * 0.7, w: '85%', h: 0.6,
    fontSize: 16, color: '333333',
  });
});

// Section slides
sections.forEach((section) => {
  const slide = pptx.addSlide();
  slide.background = { color: BRAND.accent.replace('#', '') };
  slide.addText(section, {
    x: 0.5, y: 2.5, w: '90%', h: 1.0,
    fontSize: 28, color: 'FFFFFF', bold: true, align: 'center',
  });
});

const outputName = input.fileName || 'pitchbook';
await pptx.writeFile({ fileName: `./${outputName}.pptx` });
console.log(`OK: wrote ${outputName}.pptx`);
