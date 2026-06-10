// shared/docx-report.js
// Declarative Word report composer. Each docx generator builds a `spec`
// and calls buildReport(spec) → Buffer.
//
// spec = {
//   docTitle, docSubtitle?, project?, date?, author?, lang?, confidential?,
//   includeToc?, headerLabel?,
//   sections: [
//     { heading?, level?, paragraphs?: [..], bullets?: [..],
//       facts?: [[label, value], ..], table?: { headers, rows, .. },
//       checklist?: [..], keyValues?: [[label, value]..], pageBreakAfter? }
//   ],
//   disclaimer?: string,
// }

import { Document, Packer } from 'docx';
import { BRAND } from './branding.js';
import {
  coverPage, tocPage, heading, body, bullet, spacer, pageBreak,
  defaultHeader, defaultFooter, disclaimerBlock, keyValueLine,
} from './layouts.js';
import { dataTable, factsTable, checklistTable } from './table-builder.js';

export function sectionChildren(section, lang = 'fr') {
  const children = [];
  if (section.heading) children.push(heading(section.heading, section.level || 1));
  for (const p of section.paragraphs || []) {
    if (typeof p === 'string') children.push(body(p, { align: 'justify' }));
    else children.push(body(p.text, p));
  }
  for (const kv of section.keyValues || []) {
    children.push(keyValueLine(kv[0], kv[1]));
  }
  for (const b of section.bullets || []) {
    if (typeof b === 'string') children.push(bullet(b));
    else children.push(bullet(b.text, b.level || 0));
  }
  if (section.facts?.length) {
    children.push(factsTable(section.facts));
    children.push(spacer(160));
  }
  if (section.table) {
    if (section.table.caption) children.push(body(section.table.caption, { bold: true, color: BRAND.darkBlue }));
    children.push(dataTable(section.table));
    children.push(spacer(160));
  }
  for (const t of section.tables || []) {
    if (t.caption) children.push(body(t.caption, { bold: true, color: BRAND.darkBlue }));
    children.push(dataTable(t));
    children.push(spacer(160));
  }
  if (section.checklist?.length) {
    children.push(checklistTable(section.checklist, { lang }));
    children.push(spacer(160));
  }
  for (const sub of section.subsections || []) {
    children.push(...sectionChildren({ ...sub, level: sub.level || (section.level || 1) + 1 }, lang));
  }
  if (section.pageBreakAfter) children.push(pageBreak());
  return children;
}

export async function buildReport(spec) {
  const lang = spec.lang || 'fr';
  const children = [];

  children.push(...coverPage({
    docTitle: spec.docTitle,
    docSubtitle: spec.docSubtitle,
    project: spec.project,
    date: spec.date,
    author: spec.author,
    confidential: spec.confidential !== false,
    lang,
  }));

  if (spec.includeToc !== false) children.push(...tocPage(lang));

  for (const section of spec.sections || []) {
    children.push(...sectionChildren(section, lang));
  }

  if (spec.disclaimer) {
    children.push(...disclaimerBlock(spec.disclaimer, lang === 'fr' ? 'Avertissement' : 'Important Disclaimer'));
  }

  const doc = new Document({
    creator: BRAND.brandName,
    title: spec.docTitle,
    styles: {
      default: {
        heading1: { run: { size: 30, bold: true, color: BRAND.navy, font: BRAND.font }, paragraph: { spacing: { before: 360, after: 120 } } },
        heading2: { run: { size: 24, bold: true, color: BRAND.darkBlue, font: BRAND.font }, paragraph: { spacing: { before: 240, after: 100 } } },
        heading3: { run: { size: 21, bold: true, color: BRAND.accent, font: BRAND.font }, paragraph: { spacing: { before: 200, after: 80 } } },
      },
    },
    features: { updateFields: true },
    sections: [{
      headers: { default: defaultHeader({ confidential: spec.confidential !== false, label: spec.headerLabel, lang }) },
      footers: { default: defaultFooter() },
      children,
    }],
  });

  return Packer.toBuffer(doc);
}

// ---------------------------------------------------------------------------
// Convenience: letter-style document (no TOC, no cover, letterhead block)
// ---------------------------------------------------------------------------

export async function buildLetter(spec) {
  const lang = spec.lang || 'fr';
  const children = [];

  children.push(body(BRAND.brandName, { bold: true, size: 26, color: BRAND.navy }));
  if (spec.senderLines) for (const l of spec.senderLines) children.push(body(l, { size: 18, color: BRAND.medGray }));
  children.push(spacer(240));
  if (spec.recipientLines) {
    for (const l of spec.recipientLines) children.push(body(l, { size: 20 }));
    children.push(spacer(240));
  }
  children.push(body(
    spec.date || new Date().toLocaleDateString(lang === 'fr' ? 'fr-FR' : 'en-GB', { year: 'numeric', month: 'long', day: 'numeric' }),
    { align: 'right', color: BRAND.medGray },
  ));
  children.push(spacer(160));
  if (spec.subject) children.push(body(`${lang === 'fr' ? 'Objet' : 'Re'}: ${spec.subject}`, { bold: true }));
  children.push(spacer(160));

  for (const section of spec.sections || []) {
    children.push(...sectionChildren({ ...section, level: section.level || 2 }, lang));
  }

  if (spec.signatureLines) {
    children.push(spacer(400));
    for (const l of spec.signatureLines) children.push(body(l));
  }
  if (spec.disclaimer) {
    children.push(...disclaimerBlock(spec.disclaimer, lang === 'fr' ? 'Avertissement' : 'Important Notice'));
  }

  const doc = new Document({
    creator: BRAND.brandName,
    title: spec.docTitle || spec.subject || 'Letter',
    sections: [{
      headers: { default: defaultHeader({ confidential: spec.confidential !== false, lang }) },
      footers: { default: defaultFooter() },
      children,
    }],
  });
  return Packer.toBuffer(doc);
}
