// generators/ib/leveraged-finance-memo.js
// LevFin term sheet + sizing (xlsx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildWorkbook, assumptionsSheet } from '../../shared/xlsx-model.js';
import { formatMultiple } from '../../shared/financial-helpers.js';

export const metadata = {
  id: 'leveraged-finance-memo',
  name: 'Leveraged Finance Memo (Term Sheet + Sizing)',
  vertical: 'ib',
  outputType: 'xlsx',
  estimatedPages: '4-6 sheets',
  requiredInput: ['borrower'],
  optionalInput: ['ebitda', 'tranches', 'covenants', 'assumptions', 'sensitivities'],
};

export async function generate(input = {}) {
  const borrower = input.borrower || 'Emprunteur';
  const ebitda = input.ebitda || 0;
  const tranches = input.tranches || [
    { name: 'Term Loan B', amount: ebitda * 4, rate: 'E+450', maturity: '7 ans' },
    { name: 'RCF', amount: ebitda * 0.75, rate: 'E+375', maturity: '6,5 ans' },
  ];
  const totalDebt = tranches.reduce((a, t) => a + (t.amount || 0), 0);

  const sheets = [
    {
      name: 'Term Sheet',
      sectionTitle: `Term Sheet indicatif — ${borrower}`,
      table: {
        headers: ['Tranche', 'Montant (m€)', 'Marge', 'Maturité', 'Amortissement', 'Sûretés'],
        rows: [
          ...tranches.map((t) => [t.name, t.amount ?? 0, t.rate || '—', t.maturity || '—', t.amortization || 'Bullet', t.security || 'Senior secured']),
          ['Total dette', totalDebt, '', '', '', ''],
        ],
        totalRowIndex: tranches.length,
      },
      columns: [{ width: 24 }, { width: 16 }, { width: 12 }, { width: 12 }, { width: 16 }, { width: 18 }],
    },
    {
      name: 'Leviers',
      sectionTitle: 'Analyse de levier',
      table: {
        headers: ['Métrique', 'Valeur'],
        rows: [
          ['EBITDA de référence (m€)', ebitda],
          ['Dette totale (m€)', totalDebt],
          ['Levier total', ebitda ? formatMultiple(totalDebt / ebitda) : '—'],
          ...(input.seniorDebt != null ? [['Levier senior', ebitda ? formatMultiple(input.seniorDebt / ebitda) : '—']] : []),
          ...(input.equityContribution != null ? [["Apport en fonds propres (m€)", input.equityContribution]] : []),
          ...(input.equityContribution != null ? [['% Equity', `${Math.round((input.equityContribution / (totalDebt + input.equityContribution)) * 100)}%`]] : []),
        ],
      },
      columns: [{ width: 34 }, { width: 20 }],
    },
  ];

  if (input.covenants?.length) {
    sheets.push({
      name: 'Covenants',
      sectionTitle: 'Covenants financiers',
      table: {
        headers: ['Covenant', 'Seuil', 'Test', 'Fréquence'],
        rows: input.covenants.map((c) => [c.name, c.threshold || '—', c.test || '—', c.frequency || 'Trimestrielle']),
      },
      columns: [{ width: 30 }, { width: 16 }, { width: 24 }, { width: 16 }],
    });
  }

  if (input.sensitivities?.length) {
    sheets.push({
      name: 'Sensibilités',
      sectionTitle: 'Analyse de sensibilité (couverture du service de la dette)',
      table: {
        headers: ['Scénario', ...(input.sensitivities[0]?.values || []).map((_, i) => `Année ${i + 1}`)],
        rows: input.sensitivities.map((s) => [s.name, ...(s.values || [])]),
      },
      freezeHeader: true,
    });
  }

  if (input.assumptions) {
    sheets.push(assumptionsSheet(input.assumptions));
  }

  return buildWorkbook({ title: `LevFin — ${borrower}`, sheets });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
