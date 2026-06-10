// shared/layouts.js
// Word (docx) building blocks: cover page, TOC, headings, body text,
// bullets, headers/footers with page numbers and confidentiality marks.

import {
  Paragraph, TextRun, HeadingLevel, AlignmentType, Header, Footer,
  PageNumber, TableOfContents, PageBreak, BorderStyle,
} from 'docx';
import { BRAND } from './branding.js';

export function title(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, size: 40, font: BRAND.font, color: BRAND.navy })],
    spacing: { after: 120 },
  });
}

export function subtitle(text) {
  return new Paragraph({
    children: [new TextRun({ text, size: 24, font: BRAND.font, color: BRAND.medGray, italics: true })],
    spacing: { after: 300 },
  });
}

export function heading(text, level = 1) {
  const map = { 1: HeadingLevel.HEADING_1, 2: HeadingLevel.HEADING_2, 3: HeadingLevel.HEADING_3 };
  return new Paragraph({
    text,
    heading: map[level] || HeadingLevel.HEADING_1,
    spacing: { before: level === 1 ? 360 : 240, after: 120 },
  });
}

export function body(text, options = {}) {
  return new Paragraph({
    children: [new TextRun({
      text,
      size: options.size || 20,
      font: BRAND.font,
      color: options.color || BRAND.text,
      bold: options.bold || false,
      italics: options.italic || false,
    })],
    spacing: { after: 120, line: 276 },
    alignment:
      options.align === 'center' ? AlignmentType.CENTER
        : options.align === 'right' ? AlignmentType.RIGHT
          : options.align === 'justify' ? AlignmentType.JUSTIFIED
            : AlignmentType.LEFT,
  });
}

export function bullet(text, level = 0) {
  return new Paragraph({
    children: [new TextRun({ text, size: 20, font: BRAND.font, color: BRAND.text })],
    bullet: { level },
    spacing: { after: 60, line: 276 },
  });
}

export function spacer(after = 200) {
  return new Paragraph({ text: '', spacing: { after } });
}

export function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

export function keyValueLine(label, value) {
  return new Paragraph({
    children: [
      new TextRun({ text: `${label}: `, bold: true, size: 20, font: BRAND.font, color: BRAND.navy }),
      new TextRun({ text: String(value ?? '—'), size: 20, font: BRAND.font, color: BRAND.text }),
    ],
    spacing: { after: 80 },
  });
}

/** Standard cover page children (title, subtitle, meta lines, confidentiality). */
export function coverPage({ docTitle, docSubtitle, project, date, author, confidential = true, lang = 'fr' }) {
  const children = [
    spacer(2400),
    new Paragraph({
      children: [new TextRun({ text: BRAND.brandName, bold: true, size: 28, font: BRAND.font, color: BRAND.accent })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 600 },
    }),
    new Paragraph({
      children: [new TextRun({ text: docTitle, bold: true, size: 52, font: BRAND.font, color: BRAND.navy })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
    }),
  ];
  if (docSubtitle) {
    children.push(new Paragraph({
      children: [new TextRun({ text: docSubtitle, size: 28, font: BRAND.font, color: BRAND.medGray })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 },
    }));
  }
  if (project) {
    children.push(new Paragraph({
      children: [new TextRun({ text: project, size: 24, font: BRAND.font, color: BRAND.darkBlue, italics: true })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
    }));
  }
  children.push(new Paragraph({
    children: [new TextRun({
      text: date || new Date().toLocaleDateString(lang === 'fr' ? 'fr-FR' : 'en-GB', { year: 'numeric', month: 'long', day: 'numeric' }),
      size: 22, font: BRAND.font, color: BRAND.medGray,
    })],
    alignment: AlignmentType.CENTER,
    spacing: { after: 120 },
  }));
  if (author) {
    children.push(new Paragraph({
      children: [new TextRun({ text: author, size: 20, font: BRAND.font, color: BRAND.medGray })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 120 },
    }));
  }
  if (confidential) {
    children.push(spacer(1600));
    children.push(new Paragraph({
      children: [new TextRun({
        text: lang === 'fr' ? BRAND.confidentialityFr.toUpperCase() : BRAND.confidentiality.toUpperCase(),
        bold: true, size: 20, font: BRAND.font, color: BRAND.negative,
      })],
      alignment: AlignmentType.CENTER,
      border: {
        top: { style: BorderStyle.SINGLE, size: 6, color: BRAND.negative, space: 4 },
        bottom: { style: BorderStyle.SINGLE, size: 6, color: BRAND.negative, space: 4 },
      },
    }));
  }
  children.push(pageBreak());
  return children;
}

/** TOC page children (Word refreshes field on open). */
export function tocPage(lang = 'fr') {
  return [
    heading(lang === 'fr' ? 'Sommaire' : 'Table of Contents', 1),
    new TableOfContents('TOC', { hyperlink: true, headingStyleRange: '1-3' }),
    pageBreak(),
  ];
}

export function defaultHeader({ confidential = true, label, lang = 'fr' } = {}) {
  const text = label
    || (confidential ? (lang === 'fr' ? BRAND.confidentialityFr : BRAND.confidentiality) : BRAND.brandName);
  return new Header({
    children: [new Paragraph({
      children: [new TextRun({ text, size: 16, font: BRAND.font, color: BRAND.medGray, italics: true })],
      alignment: AlignmentType.RIGHT,
    })],
  });
}

export function defaultFooter() {
  return new Footer({
    children: [new Paragraph({
      children: [
        new TextRun({ text: `${BRAND.brandName}  |  `, size: 16, font: BRAND.font, color: BRAND.medGray }),
        new TextRun({ children: [PageNumber.CURRENT], size: 16, font: BRAND.font, color: BRAND.medGray }),
      ],
      alignment: AlignmentType.CENTER,
    })],
  });
}

export function disclaimerBlock(text, langTitle = 'Avertissement') {
  return [
    spacer(300),
    heading(langTitle, 2),
    body(text, { size: 16, color: BRAND.medGray, align: 'justify', italic: true }),
  ];
}
