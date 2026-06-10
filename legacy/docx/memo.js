import { Document, Packer, Paragraph, TextRun, HeadingLevel } from 'docx';
import fs from 'node:fs';
import { BRAND } from '../../shared/branding.js';

const inputPath = process.argv[2] || './input.json';
const input = JSON.parse(fs.readFileSync(inputPath, 'utf8'));

const sections = [];

// Title
sections.push(
  new Paragraph({
    text: input.title || 'Investment Memorandum',
    heading: HeadingLevel.TITLE,
  }),
  new Paragraph({
    children: [new TextRun({ text: BRAND.confidentiality, italics: true, color: '999999' })],
  }),
  new Paragraph({ text: '' }),
);

// Sections
const memoSections = input.sections || [
  { heading: 'Executive Summary', content: '[Summary content]' },
  { heading: 'Investment Thesis', content: '[Thesis content]' },
  { heading: 'Company Overview', content: '[Company content]' },
  { heading: 'Financial Analysis', content: '[Financial content]' },
  { heading: 'Risks & Mitigants', content: '[Risk content]' },
  { heading: 'Recommendation', content: '[Recommendation]' },
];

memoSections.forEach((s) => {
  sections.push(
    new Paragraph({ text: s.heading, heading: HeadingLevel.HEADING_1 }),
    new Paragraph({ text: s.content }),
    new Paragraph({ text: '' }),
  );
});

const doc = new Document({
  creator: 'M&IA',
  sections: [{ children: sections }],
});

const buffer = await Packer.toBuffer(doc);
const outputName = input.fileName || 'memo';
fs.writeFileSync(`./${outputName}.docx`, buffer);
console.log(`OK: wrote ${outputName}.docx`);
