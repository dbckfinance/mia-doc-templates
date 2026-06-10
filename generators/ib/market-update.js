// generators/ib/market-update.js
// Point marché hebdomadaire (pptx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildDeck } from '../../shared/pptx-deck.js';

export const metadata = {
  id: 'market-update',
  name: 'Point Marché Hebdomadaire',
  vertical: 'ib',
  outputType: 'pptx',
  estimatedPages: '5-8 slides',
  requiredInput: [],
  optionalInput: ['period', 'indices', 'rates', 'deals', 'highlights', 'sectorFocus'],
};

const DEFAULT_INDICES = [
  { name: 'CAC 40', level: '7 650', change: '+1,2%' },
  { name: 'EuroStoxx 50', level: '4 980', change: '+0,8%' },
  { name: 'S&P 500', level: '5 850', change: '+1,5%' },
  { name: 'VIX', level: '14,2', change: '-0,9 pt' },
];

const DEFAULT_RATES = [
  { name: 'OAT 10 ans', level: '2,95%', change: '+5 bps' },
  { name: 'Bund 10 ans', level: '2,40%', change: '+3 bps' },
  { name: 'Euribor 3M', level: '3,15%', change: 'stable' },
  { name: 'iTraxx Crossover', level: '310', change: '-8 bps' },
];

export async function generate(input = {}) {
  const period = input.period || 'Semaine en cours';
  const indices = input.indices?.length ? input.indices : DEFAULT_INDICES;
  const rates = input.rates?.length ? input.rates : DEFAULT_RATES;
  const deals = input.deals || [];
  const highlights = input.highlights || [
    'Marchés actions en hausse portés par les résultats trimestriels',
    'Détente sur les spreads de crédit high yield',
    'Pipeline ECM actif sur les midcaps européennes',
  ];

  const slides = [
    { type: 'section', title: 'Synthèse de la semaine' },
    { type: 'content', title: 'Faits marquants', bullets: highlights },
    {
      type: 'table',
      title: 'Indices actions',
      headers: ['Indice', 'Niveau', 'Variation hebdo'],
      rows: indices.map((i) => [i.name, i.level, i.change]),
    },
    {
      type: 'table',
      title: 'Taux & crédit',
      headers: ['Instrument', 'Niveau', 'Variation hebdo'],
      rows: rates.map((r) => [r.name, r.level, r.change]),
    },
  ];

  if (deals.length) {
    slides.push({
      type: 'table',
      title: 'Transactions annoncées',
      headers: ['Cible', 'Acquéreur', 'Valeur', 'Secteur'],
      rows: deals.map((d) => [d.target || '—', d.acquirer || '—', d.value || 'n.d.', d.sector || '—']),
    });
  }

  if (input.sectorFocus) {
    slides.push({
      type: 'content',
      title: `Focus secteur — ${input.sectorFocus.name || 'Secteur'}`,
      bullets: input.sectorFocus.points || [],
    });
  }

  return buildDeck({
    title: 'Point marché',
    subtitle: period,
    confidential: false,
    slides,
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
