// generators/hedge-fund/trade-idea.js
// Pitch d'idée de trade (pptx, 3-5 slides)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildDeck } from '../../shared/pptx-deck.js';

export const metadata = {
  id: 'trade-idea',
  name: 'Idée de Trade (Pitch)',
  vertical: 'hedge-fund',
  outputType: 'pptx',
  estimatedPages: '3-5 slides',
  requiredInput: ['ticker'],
  optionalInput: ['company', 'direction', 'currentPrice', 'targetPrice', 'stopPrice', 'thesis', 'catalysts', 'risks', 'timeline'],
};

export async function generate(input = {}) {
  const ticker = input.ticker || 'TICKER';
  const company = input.company || ticker;
  const direction = (input.direction || 'long').toUpperCase();
  const current = input.currentPrice;
  const target = input.targetPrice;
  const stop = input.stopPrice;
  const upside = current && target ? `${((target / current - 1) * 100).toFixed(1)}%` : 'n.d.';
  const downside = current && stop ? `${((stop / current - 1) * 100).toFixed(1)}%` : 'n.d.';

  return buildDeck({
    title: `${direction} ${company} (${ticker})`,
    subtitle: `Idée de trade — ${new Date().toLocaleDateString('fr-FR')}`,
    confidential: true,
    slides: [
      {
        type: 'facts',
        title: 'Setup',
        facts: [
          { label: 'Direction', value: direction },
          { label: 'Cours actuel', value: current != null ? String(current) : 'n.d.' },
          { label: 'Objectif', value: target != null ? String(target) : 'n.d.' },
          { label: 'Stop', value: stop != null ? String(stop) : 'n.d.' },
          { label: 'Potentiel', value: upside },
          { label: 'Risque', value: downside },
          { label: 'Horizon', value: input.timeline || '3-6 mois' },
        ],
      },
      {
        type: 'content',
        title: 'Thèse en 3 points',
        bullets: input.thesis || [
          'Dislocation de valorisation vs fondamentaux',
          'Catalyseur identifié à court terme',
          'Asymétrie risque/rendement favorable (>2:1)',
        ],
      },
      {
        type: 'content',
        title: 'Catalyseurs et timeline',
        bullets: input.catalysts || [
          'Résultats trimestriels — date à confirmer',
          'Décision réglementaire attendue',
          'Flux : fin de la pression vendeuse technique',
        ],
      },
      {
        type: 'content',
        title: 'Risques',
        bullets: input.risks || [
          'Invalidation de la thèse si dégradation des marges',
          'Risque de marché global (beta)',
          'Liquidité du titre en cas de sortie rapide',
        ],
      },
    ],
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
