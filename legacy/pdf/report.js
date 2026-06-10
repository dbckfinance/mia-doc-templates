import PDFDocument from 'pdfkit';
import fs from 'node:fs';
import { BRAND } from '../../shared/branding.js';

const inputPath = process.argv[2] || './input.json';
const input = JSON.parse(fs.readFileSync(inputPath, 'utf8'));

const doc = new PDFDocument({ size: 'A4', margin: 50 });
const outputName = input.fileName || 'report';
const stream = fs.createWriteStream(`./${outputName}.pdf`);
doc.pipe(stream);

// Title page
doc.fontSize(28).fillColor(BRAND.primary)
  .text(input.title || 'Research Report', { align: 'center' });
doc.moveDown(2);
doc.fontSize(14).fillColor('#666666')
  .text(BRAND.confidentiality, { align: 'center' });
doc.moveDown(1);
doc.fontSize(12).fillColor('#333333')
  .text(input.subtitle || '', { align: 'center' });

doc.addPage();

// Sections
const sections = input.sections || [
  { heading: 'Executive Summary', content: '[Summary]' },
  { heading: 'Market Analysis', content: '[Analysis]' },
  { heading: 'Findings', content: '[Findings]' },
  { heading: 'Conclusion', content: '[Conclusion]' },
];

sections.forEach((s) => {
  doc.fontSize(18).fillColor(BRAND.primary).text(s.heading);
  doc.moveDown(0.5);
  doc.fontSize(11).fillColor('#333333').text(s.content);
  doc.moveDown(1.5);
});

doc.end();
stream.on('finish', () => {
  console.log(`OK: wrote ${outputName}.pdf`);
});
