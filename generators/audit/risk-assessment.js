// generators/audit/risk-assessment.js
// Matrice des risques + heatmap (xlsx + pptx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildWorkbook } from '../../shared/xlsx-model.js';
import { buildDeck } from '../../shared/pptx-deck.js';

export const metadata = {
  id: 'risk-assessment',
  name: 'Cartographie des Risques',
  vertical: 'audit',
  outputType: 'multi',
  outputs: ['xlsx', 'pptx'],
  estimatedPages: '2 sheets + 3 slides',
  requiredInput: ['company'],
  optionalInput: ['risks'],
};

const DEFAULT_RISKS = [
  { name: 'Cyber-attaque / ransomware', category: 'IT', probability: 4, impact: 5, mitigation: 'Plan de continuité, EDR, sauvegardes hors-ligne', owner: 'DSI' },
  { name: 'Défaillance client majeur', category: 'Crédit', probability: 3, impact: 4, mitigation: 'Assurance-crédit, suivi des encours', owner: 'DAF' },
  { name: 'Non-conformité RGPD', category: 'Réglementaire', probability: 2, impact: 4, mitigation: 'DPO, registre des traitements, audits annuels', owner: 'Juridique' },
  { name: 'Rupture supply chain', category: 'Opérationnel', probability: 3, impact: 3, mitigation: 'Double sourcing, stocks de sécurité', owner: 'Achats' },
  { name: 'Départ d\'hommes clés', category: 'RH', probability: 3, impact: 3, mitigation: 'Plans de succession, rétention', owner: 'DRH' },
];

function level(score) {
  if (score >= 16) return 'Critique';
  if (score >= 9) return 'Élevé';
  if (score >= 4) return 'Modéré';
  return 'Faible';
}

export async function generate(input = {}) {
  const company = input.company || 'Société';
  const risks = (input.risks?.length ? input.risks : DEFAULT_RISKS).map((r) => ({
    ...r,
    score: (r.probability || 1) * (r.impact || 1),
  })).sort((a, b) => b.score - a.score);

  const xlsx = await buildWorkbook({
    title: `Cartographie des risques — ${company}`,
    sheets: [
      {
        name: 'Registre des risques',
        sectionTitle: `Registre des risques — ${company}`,
        table: {
          headers: ['Risque', 'Catégorie', 'Probabilité (1-5)', 'Impact (1-5)', 'Score', 'Niveau', 'Mitigation', 'Responsable'],
          rows: risks.map((r) => [r.name, r.category || '—', r.probability, r.impact, r.score, level(r.score), r.mitigation || '—', r.owner || '—']),
        },
        columns: [{ width: 34 }, { width: 16 }, { width: 16 }, { width: 12 }, { width: 10 }, { width: 12 }, { width: 44 }, { width: 14 }],
        freezeHeader: true,
      },
      {
        name: 'Synthèse',
        sectionTitle: 'Synthèse par niveau',
        table: {
          headers: ['Niveau', 'Nombre de risques'],
          rows: ['Critique', 'Élevé', 'Modéré', 'Faible'].map((lvl) => [lvl, risks.filter((r) => level(r.score) === lvl).length]),
        },
        columns: [{ width: 16 }, { width: 20 }],
      },
    ],
  });

  const pptx = await buildDeck({
    title: `Cartographie des risques — ${company}`,
    subtitle: 'Évaluation probabilité x impact',
    confidential: true,
    slides: [
      {
        type: 'chart',
        title: 'Heatmap des risques',
        chart: {
          kind: 'riskHeatmap',
          risks: risks.map((r) => ({ label: r.name, probability: r.probability, impact: r.impact })),
        },
      },
      {
        type: 'table',
        title: 'Top risques',
        headers: ['Risque', 'Score', 'Niveau', 'Mitigation'],
        rows: risks.slice(0, 8).map((r) => [r.name, String(r.score), level(r.score), r.mitigation || '—']),
      },
      {
        type: 'content',
        title: 'Prochaines étapes',
        bullets: [
          'Valider la cartographie avec le comité des risques',
          'Affecter les plans de mitigation aux responsables identifiés',
          'Revue trimestrielle des risques critiques et élevés',
        ],
      },
    ],
  });

  return [
    { fileName: 'risk-assessment.xlsx', buffer: xlsx, ext: 'xlsx' },
    { fileName: 'risk-assessment.pptx', buffer: pptx, ext: 'pptx' },
  ];
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
