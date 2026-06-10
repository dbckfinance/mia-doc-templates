// generators/hedge-fund/investment-memo.js
// Mémo de thèse d'investissement (docx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildReport } from '../../shared/docx-report.js';

export const metadata = {
  id: 'investment-memo',
  name: "Mémo de Thèse d'Investissement",
  vertical: 'hedge-fund',
  outputType: 'docx',
  estimatedPages: '8-15',
  requiredInput: ['ticker'],
  optionalInput: ['company', 'direction', 'targetPrice', 'currentPrice', 'horizon', 'thesis', 'catalysts', 'risks', 'valuation', 'positionSize'],
};

export async function generate(input = {}) {
  const ticker = input.ticker || 'TICKER';
  const company = input.company || ticker;
  const direction = (input.direction || 'long').toLowerCase();
  const current = input.currentPrice;
  const target = input.targetPrice;
  const upside = current && target ? ((target / current - 1) * 100).toFixed(1) + '%' : 'n.d.';

  return buildReport({
    docTitle: `Thèse d'investissement — ${company} (${ticker})`,
    docSubtitle: `${direction === 'short' ? 'SHORT' : 'LONG'} — Horizon ${input.horizon || '12-18 mois'}`,
    vertical: 'hedge-fund',
    docId: 'investment-memo',
    confidential: true,
    sections: [
      {
        heading: 'Résumé de la position',
        blocks: [
          { type: 'facts', facts: [
            { label: 'Ticker', value: ticker },
            { label: 'Direction', value: direction === 'short' ? 'Short' : 'Long' },
            { label: 'Cours actuel', value: current != null ? String(current) : 'n.d.' },
            { label: 'Objectif de cours', value: target != null ? String(target) : 'n.d.' },
            { label: 'Potentiel', value: upside },
            { label: 'Taille de position', value: input.positionSize || '2-4% NAV' },
            { label: 'Horizon', value: input.horizon || '12-18 mois' },
          ] },
        ],
      },
      {
        heading: 'Thèse d\'investissement',
        blocks: [
          { type: 'bullets', items: input.thesis || [
            'Valorisation décotée par rapport aux pairs malgré des fondamentaux supérieurs',
            'Inflexion de marge attendue non reflétée dans le consensus',
            'Allocation du capital disciplinée (rachats d\'actions, désendettement)',
          ] },
        ],
      },
      {
        heading: 'Catalyseurs',
        blocks: [
          { type: 'bullets', items: input.catalysts || [
            'Publication des résultats trimestriels (beat & raise attendu)',
            'Cession d\'actifs non stratégiques',
            'Journée investisseurs avec nouveaux objectifs moyen terme',
          ] },
        ],
      },
      {
        heading: 'Valorisation',
        blocks: input.valuation?.rows?.length
          ? [{ type: 'table', headers: input.valuation.headers || ['Méthode', 'Valeur'], rows: input.valuation.rows }]
          : [{ type: 'p', text: 'Le multiple cible appliqué aux estimations à 12 mois conduit à l\'objectif de cours retenu. Le scénario défavorable et le scénario optimiste encadrent la valeur centrale (détail en annexe du modèle).' }],
      },
      {
        heading: 'Risques et points de surveillance',
        blocks: [
          { type: 'bullets', items: input.risks || [
            'Sensibilité macro (consommation, taux)',
            'Risque d\'exécution sur le plan stratégique',
            'Pression concurrentielle sur les prix',
          ] },
        ],
      },
      {
        heading: 'Gestion de la position',
        blocks: [
          { type: 'kv', label: 'Stop / réévaluation', value: input.stopLoss || 'Réévaluation complète si -15% vs prix d\'entrée' },
          { type: 'kv', label: 'Plan de sortie', value: input.exitPlan || 'Allègement progressif à l\'approche de l\'objectif de cours' },
        ],
      },
    ],
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
