// generators/ib/credit-analysis.js
// Analyse crédit : métriques et covenants (xlsx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildWorkbook } from '../../shared/xlsx-model.js';
import { creditMetrics, formatMultiple, formatPercent } from '../../shared/financial-helpers.js';

export const metadata = {
  id: 'credit-analysis',
  name: 'Analyse Crédit',
  vertical: 'ib',
  outputType: 'xlsx',
  estimatedPages: '3 sheets',
  requiredInput: ['company'],
  optionalInput: ['years', 'covenants'],
};

const DEFAULT_YEARS = [
  { year: 'A-2', ebitda: 80, totalDebt: 320, cash: 30, interestExpense: 18, fcf: 35 },
  { year: 'A-1', ebitda: 88, totalDebt: 300, cash: 35, interestExpense: 17, fcf: 42 },
  { year: 'A', ebitda: 95, totalDebt: 280, cash: 40, interestExpense: 16, fcf: 50 },
];

const DEFAULT_COVENANTS = [
  { name: 'Levier net maximum', threshold: '4,0x', basis: 'Dette nette / EBITDA' },
  { name: 'Couverture des intérêts minimum', threshold: '3,0x', basis: 'EBITDA / Intérêts' },
  { name: 'Capex maximum', threshold: '30 m€/an', basis: 'Capex cumulé' },
];

export async function generate(input = {}) {
  const company = input.company || 'Société';
  const years = input.years?.length ? input.years : DEFAULT_YEARS;
  const covenants = input.covenants?.length ? input.covenants : DEFAULT_COVENANTS;

  const metricRows = years.map((y) => {
    const m = creditMetrics(y);
    return [
      y.year,
      y.ebitda,
      y.totalDebt,
      y.cash,
      m.netDebt,
      m.leverage != null ? formatMultiple(m.leverage, { decimals: 2 }) : '—',
      m.interestCoverage != null ? formatMultiple(m.interestCoverage, { decimals: 2 }) : '—',
      m.fcfToDebt != null ? formatPercent(m.fcfToDebt) : '—',
    ];
  });

  const latest = creditMetrics(years[years.length - 1]);

  return buildWorkbook({
    title: `Analyse crédit — ${company}`,
    sheets: [
      {
        name: 'Métriques',
        sectionTitle: `Métriques de crédit — ${company}`,
        table: {
          headers: ['Année', 'EBITDA', 'Dette brute', 'Trésorerie', 'Dette nette', 'Levier net', 'Couverture intérêts', 'FCF / Dette'],
          rows: metricRows,
        },
        columns: [{ width: 10 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 18 }, { width: 12 }],
        freezeHeader: true,
      },
      {
        name: 'Covenants',
        sectionTitle: 'Covenants bancaires',
        table: {
          headers: ['Covenant', 'Seuil', 'Base de calcul', 'Niveau actuel', 'Headroom'],
          rows: covenants.map((c) => {
            let current = '—';
            let headroom = 'À évaluer';
            if (/levier/i.test(c.name) && latest.leverage != null) {
              current = formatMultiple(latest.leverage, { decimals: 2 });
              const thr = parseFloat(String(c.threshold).replace(',', '.'));
              if (Number.isFinite(thr)) headroom = formatPercent(1 - latest.leverage / thr);
            }
            if (/couverture/i.test(c.name) && latest.interestCoverage != null) {
              current = formatMultiple(latest.interestCoverage, { decimals: 2 });
              const thr = parseFloat(String(c.threshold).replace(',', '.'));
              if (Number.isFinite(thr)) headroom = formatPercent(latest.interestCoverage / thr - 1);
            }
            return [c.name, c.threshold, c.basis, current, headroom];
          }),
        },
        columns: [{ width: 30 }, { width: 14 }, { width: 26 }, { width: 14 }, { width: 12 }],
      },
      {
        name: 'Synthèse',
        sectionTitle: 'Synthèse crédit',
        table: {
          headers: ['Indicateur', 'Valeur'],
          rows: [
            ['Dette nette actuelle (m€)', latest.netDebt ?? '—'],
            ['Levier net', latest.leverage != null ? formatMultiple(latest.leverage, { decimals: 2 }) : '—'],
            ['Couverture des intérêts', latest.interestCoverage != null ? formatMultiple(latest.interestCoverage, { decimals: 2 }) : '—'],
            ['Profil', latest.leverage != null && latest.leverage < 3 ? 'Investment grade implicite' : 'Levier élevé — surveiller'],
          ],
        },
        columns: [{ width: 32 }, { width: 26 }],
      },
    ],
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
