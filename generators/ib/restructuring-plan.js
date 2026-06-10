// generators/ib/restructuring-plan.js
// Plan de restructuration (docx + xlsx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildReport } from '../../shared/docx-report.js';
import { buildWorkbook } from '../../shared/xlsx-model.js';
import { disclaimerFor } from '../../shared/disclaimer-library.js';

export const metadata = {
  id: 'restructuring-plan',
  name: 'Plan de Restructuration',
  vertical: 'ib',
  outputType: 'docx+xlsx',
  estimatedPages: '15-25 pages + modèle',
  requiredInput: ['company'],
  optionalInput: ['situation', 'debtStructure', 'options', 'recommendation', 'cashflow', 'stakeholders'],
};

export async function generate(input = {}) {
  const company = input.company || 'Société';

  const sections = [
    {
      heading: 'Diagnostic de la situation',
      paragraphs: input.situation ? [input.situation] : ['Diagnostic à compléter.'],
      bullets: input.causes || [],
      pageBreakAfter: false,
    },
  ];

  if (input.debtStructure?.length) {
    sections.push({
      heading: "Structure d'endettement actuelle",
      table: {
        headers: ['Instrument', 'Encours (m€)', 'Maturité', 'Taux', 'Sûretés', 'Créancier'],
        rows: input.debtStructure.map((d) => [d.instrument, d.amount ?? '—', d.maturity || '—', d.rate || '—', d.security || '—', d.lender || '—']),
      },
    });
  }

  if (input.options?.length) {
    sections.push({
      heading: 'Options de restructuration',
      table: {
        headers: ['Option', 'Description', 'Avantages', 'Risques'],
        rows: input.options.map((o) => [o.name, o.description || '—', o.pros || '—', o.cons || '—']),
      },
    });
  }

  sections.push({
    heading: 'Recommandation',
    paragraphs: [input.recommendation || 'Recommandation à formuler après analyse des options.'],
  });

  if (input.stakeholders?.length) {
    sections.push({
      heading: 'Cartographie des parties prenantes',
      table: {
        headers: ['Partie prenante', 'Exposition', 'Position attendue', 'Stratégie de négociation'],
        rows: input.stakeholders.map((s) => [s.name, s.exposure || '—', s.stance || '—', s.strategy || '—']),
      },
    });
  }

  if (input.timeline?.length) {
    sections.push({ heading: 'Calendrier de mise en œuvre', bullets: input.timeline });
  }

  const docBuffer = await buildReport({
    docTitle: 'Plan de Restructuration',
    docSubtitle: company,
    project: input.project,
    date: input.date,
    lang: 'fr',
    sections,
    disclaimer: disclaimerFor('ib', metadata.id),
  });

  const cf = input.cashflow || {};
  const wbBuffer = await buildWorkbook({
    title: `Restructuration — ${company}`,
    sheets: [
      {
        name: 'Dette',
        sectionTitle: "Structure d'endettement",
        table: {
          headers: ['Instrument', 'Encours (m€)', 'Maturité', 'Taux', 'Sûretés'],
          rows: (input.debtStructure || []).map((d) => [d.instrument, d.amount ?? 0, d.maturity || '—', d.rate || '—', d.security || '—']),
        },
        columns: [{ width: 30 }, { width: 16 }, { width: 14 }, { width: 12 }, { width: 20 }],
      },
      {
        name: 'Trésorerie 13 semaines',
        sectionTitle: 'Prévision de trésorerie 13 semaines',
        table: {
          headers: ['Ligne', ...(cf.weeks || Array.from({ length: 13 }, (_, i) => `S${i + 1}`))],
          rows: (cf.rows || [
            { name: 'Encaissements', values: Array(13).fill(0) },
            { name: 'Décaissements', values: Array(13).fill(0) },
            { name: 'Flux net', values: Array(13).fill(0) },
            { name: 'Trésorerie fin de semaine', values: Array(13).fill(0) },
          ]).map((r) => [r.name, ...(r.values || [])]),
        },
        freezeHeader: true,
      },
    ],
  });

  return [
    { buffer: docBuffer, ext: 'docx', suffix: 'plan' },
    { buffer: wbBuffer, ext: 'xlsx', suffix: 'modele' },
  ];
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
