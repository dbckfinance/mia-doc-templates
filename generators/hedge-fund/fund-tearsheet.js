// generators/hedge-fund/fund-tearsheet.js
// Factsheet une page du fonds (pptx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildDeck } from '../../shared/pptx-deck.js';
import { formatPercent } from '../../shared/financial-helpers.js';

export const metadata = {
  id: 'fund-tearsheet',
  name: 'Fund Tearsheet (Factsheet)',
  vertical: 'hedge-fund',
  outputType: 'pptx',
  estimatedPages: '1-2 slides',
  requiredInput: ['fund'],
  optionalInput: ['strategy', 'aum', 'inception', 'performance', 'stats', 'terms', 'monthlyReturns'],
};

export async function generate(input = {}) {
  const fund = input.fund || 'Fonds';
  const perf = input.performance || { ytd: 0.078, oneYear: 0.124, threeYearAnn: 0.092, sinceInceptionAnn: 0.105 };
  const stats = input.stats || { sharpe: 1.35, sortino: 1.9, maxDrawdown: -0.085, volatility: 0.082, correlation: 0.45 };
  const terms = input.terms || {
    managementFee: '1,5%',
    performanceFee: '20% (high water mark)',
    liquidity: 'Mensuelle, préavis 30 jours',
    minimum: '1 m€',
  };

  return buildDeck({
    title: fund,
    subtitle: `${input.strategy || 'Long/Short Equity'} — Factsheet`,
    confidential: true,
    slides: [
      {
        type: 'facts',
        title: 'Performance et profil',
        facts: [
          { label: 'YTD', value: formatPercent(perf.ytd) },
          { label: '1 an', value: formatPercent(perf.oneYear) },
          { label: '3 ans (ann.)', value: formatPercent(perf.threeYearAnn) },
          { label: 'Depuis création (ann.)', value: formatPercent(perf.sinceInceptionAnn) },
          { label: 'Sharpe', value: String(stats.sharpe) },
          { label: 'Volatilité', value: formatPercent(stats.volatility) },
          { label: 'Drawdown max', value: formatPercent(stats.maxDrawdown) },
          { label: 'Encours', value: input.aum || 'n.d.' },
        ],
      },
      {
        type: 'table',
        title: 'Conditions',
        headers: ['Terme', 'Détail'],
        rows: [
          ['Stratégie', input.strategy || 'Long/Short Equity'],
          ['Date de création', input.inception || 'n.d.'],
          ['Frais de gestion', terms.managementFee],
          ['Commission de performance', terms.performanceFee],
          ['Liquidité', terms.liquidity],
          ['Investissement minimum', terms.minimum],
        ],
      },
    ],
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
