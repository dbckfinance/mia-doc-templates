// generators/ib/football-field.js
// Football field de valorisation (pptx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildDeck } from '../../shared/pptx-deck.js';

export const metadata = {
  id: 'football-field',
  name: 'Football Field de Valorisation',
  vertical: 'ib',
  outputType: 'pptx',
  estimatedPages: '2-3 slides',
  requiredInput: ['target'],
  optionalInput: ['methods', 'currency', 'selectedRange'],
};

const DEFAULT_METHODS = [
  { label: 'DCF (WACC ±0,5%)', low: 420, high: 520 },
  { label: 'Comparables boursiers', low: 380, high: 470 },
  { label: 'Transactions précédentes', low: 440, high: 560 },
  { label: 'LBO (TRI 20-25%)', low: 360, high: 450 },
];

export async function generate(input = {}) {
  const target = input.target || 'Cible';
  const currency = input.currency || 'm€';
  const methods = input.methods?.length ? input.methods : DEFAULT_METHODS;
  const lows = methods.map((m) => m.low);
  const highs = methods.map((m) => m.high);
  const globalLow = Math.min(...lows);
  const globalHigh = Math.max(...highs);
  const selected = input.selectedRange || {
    low: Math.round((globalLow + globalHigh) / 2 * 0.9),
    high: Math.round((globalLow + globalHigh) / 2 * 1.1),
  };

  return buildDeck({
    title: `Synthèse de valorisation — ${target}`,
    subtitle: `Football field (${currency})`,
    confidential: true,
    slides: [
      {
        type: 'chart',
        title: `Fourchettes de valorisation (${currency})`,
        chart: {
          kind: 'footballField',
          items: methods.map((m) => ({ label: m.label, low: m.low, high: m.high })),
        },
      },
      {
        type: 'table',
        title: 'Détail par méthode',
        headers: ['Méthode', `Bas (${currency})`, `Haut (${currency})`, `Milieu (${currency})`],
        rows: methods.map((m) => [m.label, m.low, m.high, Math.round((m.low + m.high) / 2)]),
      },
      {
        type: 'facts',
        title: 'Fourchette retenue',
        facts: [
          { label: 'Borne basse', value: `${selected.low} ${currency}` },
          { label: 'Borne haute', value: `${selected.high} ${currency}` },
          { label: 'Valeur centrale', value: `${Math.round((selected.low + selected.high) / 2)} ${currency}` },
        ],
      },
    ],
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
