// generators/ma/buyer-list.js
// Univers d'acheteurs (xlsx) — stratégiques + sponsors financiers

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildWorkbook } from '../../shared/xlsx-model.js';

export const metadata = {
  id: 'buyer-list',
  name: "Univers d'Acheteurs (Buyer List)",
  vertical: 'ma',
  outputType: 'xlsx',
  estimatedPages: '2-4 sheets',
  requiredInput: ['buyers'],
  optionalInput: ['project', 'criteria'],
};

const BUYER_HEADERS = [
  'Société', 'Type', 'Pays', 'CA (m€)', 'Rationale stratégique',
  'Capacité financière', 'Probabilité', 'Contact', 'Statut', 'Commentaires',
];

function buyerRow(b) {
  return [
    b.name || '—',
    b.type || '—',
    b.country || '—',
    b.revenue ?? '—',
    b.rationale || '—',
    b.capacity || '—',
    b.likelihood || '—',
    b.contact || '—',
    b.status || 'À contacter',
    b.notes || '',
  ];
}

export async function generate(input = {}) {
  const buyers = input.buyers || [];
  const strategic = buyers.filter((b) => (b.type || '').toLowerCase().includes('strat') || b.category === 'strategic');
  const financial = buyers.filter((b) => /sponsor|financ|pe|fonds/i.test(b.type || '') || b.category === 'financial');
  const others = buyers.filter((b) => !strategic.includes(b) && !financial.includes(b));

  const sheets = [];

  if (input.criteria?.length) {
    sheets.push({
      name: 'Critères',
      sectionTitle: `Critères de sélection — ${input.project || 'Projet'}`,
      table: { headers: ['Critère', 'Description'], rows: input.criteria.map((c) => [c.name || c, c.description || '']) },
      columns: [{ width: 35 }, { width: 70 }],
    });
  }

  sheets.push({
    name: 'Synthèse',
    sectionTitle: `Univers d'acheteurs — ${input.project || 'Projet'}`,
    table: {
      headers: ['Catégorie', 'Nombre', 'Dont prioritaires'],
      rows: [
        ['Acquéreurs stratégiques', strategic.length, strategic.filter((b) => /haute|high|1/i.test(String(b.likelihood))).length],
        ['Sponsors financiers', financial.length, financial.filter((b) => /haute|high|1/i.test(String(b.likelihood))).length],
        ['Autres', others.length, 0],
        ['Total', buyers.length, ''],
      ],
      totalRowIndex: 3,
    },
    columns: [{ width: 30 }, { width: 14 }, { width: 18 }],
  });

  if (strategic.length) {
    sheets.push({
      name: 'Stratégiques',
      sectionTitle: 'Acquéreurs stratégiques',
      table: { headers: BUYER_HEADERS, rows: strategic.map(buyerRow) },
      columns: [{ width: 28 }, { width: 14 }, { width: 10 }, { width: 10 }, { width: 45 }, { width: 16 }, { width: 12 }, { width: 22 }, { width: 14 }, { width: 30 }],
      freezeHeader: true,
    });
  }
  if (financial.length) {
    sheets.push({
      name: 'Sponsors',
      sectionTitle: 'Sponsors financiers',
      table: { headers: BUYER_HEADERS, rows: financial.map(buyerRow) },
      columns: [{ width: 28 }, { width: 14 }, { width: 10 }, { width: 10 }, { width: 45 }, { width: 16 }, { width: 12 }, { width: 22 }, { width: 14 }, { width: 30 }],
      freezeHeader: true,
    });
  }
  if (others.length || (!strategic.length && !financial.length)) {
    sheets.push({
      name: 'Longlist',
      sectionTitle: 'Longlist complète',
      table: { headers: BUYER_HEADERS, rows: (others.length ? others : buyers).map(buyerRow) },
      columns: [{ width: 28 }, { width: 14 }, { width: 10 }, { width: 10 }, { width: 45 }, { width: 16 }, { width: 12 }, { width: 22 }, { width: 14 }, { width: 30 }],
      freezeHeader: true,
    });
  }

  return buildWorkbook({ title: `Buyer List — ${input.project || 'Projet'}`, sheets });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
