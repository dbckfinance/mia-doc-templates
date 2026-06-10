// generators/equity-research/esg-scorecard.js
// Scorecard ESG société (pptx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildDeck } from '../../shared/pptx-deck.js';

export const metadata = {
  id: 'esg-scorecard',
  name: 'Scorecard ESG Société',
  vertical: 'equity-research',
  outputType: 'pptx',
  estimatedPages: '4-6 slides',
  requiredInput: ['ticker'],
  optionalInput: ['company', 'scores', 'peers', 'controversies', 'engagement'],
};

const DEFAULT_SCORES = [
  { criterion: 'Émissions carbone (intensité)', pillar: 'E', score: 72, trend: '↑' },
  { criterion: 'Gestion de l\'eau et déchets', pillar: 'E', score: 65, trend: '→' },
  { criterion: 'Santé & sécurité', pillar: 'S', score: 80, trend: '↑' },
  { criterion: 'Capital humain & diversité', pillar: 'S', score: 70, trend: '↑' },
  { criterion: 'Indépendance du conseil', pillar: 'G', score: 75, trend: '→' },
  { criterion: 'Rémunération alignée', pillar: 'G', score: 60, trend: '↓' },
];

export async function generate(input = {}) {
  const ticker = input.ticker || 'TICKER';
  const company = input.company || ticker;
  const scores = input.scores?.length ? input.scores : DEFAULT_SCORES;

  const pillarAvg = (p) => {
    const items = scores.filter((s) => s.pillar === p);
    return items.length ? Math.round(items.reduce((sum, s) => sum + s.score, 0) / items.length) : 0;
  };
  const e = pillarAvg('E'), s = pillarAvg('S'), g = pillarAvg('G');
  const overall = Math.round((e + s + g) / 3);

  return buildDeck({
    title: `Scorecard ESG — ${company}`,
    subtitle: ticker,
    confidential: false,
    slides: [
      {
        type: 'facts',
        title: 'Scores ESG',
        facts: [
          { label: 'Score global', value: `${overall}/100` },
          { label: 'Environnement', value: `${e}/100` },
          { label: 'Social', value: `${s}/100` },
          { label: 'Gouvernance', value: `${g}/100` },
        ],
      },
      {
        type: 'table',
        title: 'Détail par critère',
        headers: ['Critère', 'Pilier', 'Score', 'Tendance'],
        rows: scores.map((sc) => [sc.criterion, sc.pillar, `${sc.score}/100`, sc.trend || '—']),
      },
      {
        type: 'chart',
        title: 'Scores par pilier',
        chart: {
          kind: 'bar',
          categories: ['Environnement', 'Social', 'Gouvernance'],
          series: [
            { name: company, values: [e, s, g] },
            ...(input.peers?.length ? [{ name: 'Médiane pairs', values: input.peers }] : []),
          ],
        },
      },
      {
        type: 'content',
        title: 'Controverses et engagement',
        bullets: [
          ...(input.controversies || ['Aucune controverse majeure recensée sur les 24 derniers mois']),
          ...(input.engagement || ['Dialogue actionnarial en cours sur la politique de rémunération']),
        ],
      },
    ],
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
