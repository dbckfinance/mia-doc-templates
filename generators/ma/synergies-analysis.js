// generators/ma/synergies-analysis.js
// Analyse des synergies — bridge xlsx + deck pptx (multi-format)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildWorkbook } from '../../shared/xlsx-model.js';
import { buildDeck } from '../../shared/pptx-deck.js';
import { disclaimerFor } from '../../shared/disclaimer-library.js';

export const metadata = {
  id: 'synergies-analysis',
  name: 'Analyse des Synergies',
  vertical: 'ma',
  outputType: 'xlsx+pptx',
  estimatedPages: '3 sheets + 5 slides',
  requiredInput: ['synergies'],
  optionalInput: ['project', 'baseline', 'costsToAchieve', 'phasing'],
};

export async function generate(input = {}) {
  const synergies = input.synergies || [];
  const revenue = synergies.filter((s) => /revenu|revenue|cross/i.test(s.type || ''));
  const costs = synergies.filter((s) => !revenue.includes(s));
  const total = synergies.reduce((a, s) => a + (s.value || 0), 0);
  const cta = input.costsToAchieve || 0;

  // --- Workbook ---
  const detailHeaders = ['Synergie', 'Type', 'Valeur annuelle (m€)', 'Année pleine', 'Probabilité', 'Commentaires'];
  const detailRow = (s) => [s.name, s.type || '—', s.value ?? 0, s.fullYear || 'N+2', s.probability || '—', s.notes || ''];

  const wbBuffer = await buildWorkbook({
    title: `Analyse des synergies — ${input.project || 'Projet'}`,
    sheets: [
      {
        name: 'Synthèse',
        sectionTitle: 'Synthèse des synergies',
        table: {
          headers: ['Catégorie', 'Valeur annuelle (m€)', 'Nombre'],
          rows: [
            ['Synergies de coûts', costs.reduce((a, s) => a + (s.value || 0), 0), costs.length],
            ['Synergies de revenus', revenue.reduce((a, s) => a + (s.value || 0), 0), revenue.length],
            ['Coûts de mise en œuvre (one-off)', -Math.abs(cta), ''],
            ['Total net (run-rate)', total - Math.abs(cta), synergies.length],
          ],
          totalRowIndex: 3,
        },
        columns: [{ width: 36 }, { width: 24 }, { width: 12 }],
      },
      {
        name: 'Détail',
        sectionTitle: 'Détail des synergies identifiées',
        table: { headers: detailHeaders, rows: synergies.map(detailRow) },
        columns: [{ width: 36 }, { width: 16 }, { width: 20 }, { width: 14 }, { width: 14 }, { width: 40 }],
        freezeHeader: true,
      },
      ...(input.phasing ? [{
        name: 'Montée en charge',
        sectionTitle: 'Phasing de réalisation',
        table: {
          headers: ['Synergie', ...(input.phasing.years || ['N', 'N+1', 'N+2']).map(String)],
          rows: (input.phasing.rows || []).map((r) => [r.name, ...(r.values || [])]),
        },
        freezeHeader: true,
      }] : []),
    ],
  });

  // --- Deck ---
  const steps = [
    { label: 'EBITDA standalone', value: input.baseline || 0, isTotal: true },
    ...costs.map((s) => ({ label: s.name, value: s.value || 0 })),
    ...revenue.map((s) => ({ label: s.name, value: s.value || 0 })),
    ...(cta ? [{ label: 'Coûts de mise en œuvre', value: -Math.abs(cta) }] : []),
    { label: 'EBITDA combiné (run-rate)', isTotal: true },
  ];

  const deckBuffer = await buildDeck({
    deckTitle: 'Analyse des Synergies',
    deckSubtitle: input.project,
    date: input.date,
    lang: 'fr',
    slides: [
      {
        type: 'chart',
        title: 'Pont de création de valeur (m€)',
        chart: { kind: 'waterfall', steps },
      },
      {
        type: 'table',
        title: 'Détail des synergies',
        table: {
          headers: ['Synergie', 'Type', 'Valeur (m€)', 'Année pleine'],
          rows: synergies.map((s) => [s.name, s.type || '—', s.value ?? 0, s.fullYear || 'N+2']),
        },
      },
      ...(input.risks?.length ? [{ type: 'content', title: 'Risques d\'exécution', bullets: input.risks }] : []),
    ],
    disclaimer: disclaimerFor('ma', metadata.id),
  });

  return [
    { buffer: wbBuffer, ext: 'xlsx', suffix: 'detail' },
    { buffer: deckBuffer, ext: 'pptx', suffix: 'synthese' },
  ];
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
