// generators/ib/lbo-model.js
// Modèle LBO avec retours sponsor (xlsx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildWorkbook, FMT } from '../../shared/xlsx-model.js';
import { irr, moic, formatMultiple, formatPercent } from '../../shared/financial-helpers.js';

export const metadata = {
  id: 'lbo-model',
  name: 'Modèle LBO',
  vertical: 'ib',
  outputType: 'xlsx',
  estimatedPages: '5-7 sheets',
  requiredInput: ['target'],
  optionalInput: ['entryEbitda', 'entryMultiple', 'exitMultiple', 'holdYears', 'debtPct', 'ebitdaGrowth', 'interestRate', 'cashSweepPct', 'taxRate', 'transactionFeesPct'],
};

export async function generate(input = {}) {
  const target = input.target || 'Cible';
  const a = {
    entryEbitda: input.entryEbitda ?? 50,
    entryMultiple: input.entryMultiple ?? 9,
    exitMultiple: input.exitMultiple ?? input.entryMultiple ?? 9,
    holdYears: input.holdYears ?? 5,
    debtPct: input.debtPct ?? 0.55,
    ebitdaGrowth: input.ebitdaGrowth ?? 0.06,
    interestRate: input.interestRate ?? 0.07,
    cashSweepPct: input.cashSweepPct ?? 0.75,
    taxRate: input.taxRate ?? 0.25,
    capexPctOfEbitda: input.capexPctOfEbitda ?? 0.25,
    transactionFeesPct: input.transactionFeesPct ?? 0.02,
  };

  // --- Simulation ---
  const ev = a.entryEbitda * a.entryMultiple;
  const fees = ev * a.transactionFeesPct;
  const debt0 = ev * a.debtPct;
  const equity0 = ev + fees - debt0;

  let ebitda = a.entryEbitda;
  let debt = debt0;
  const schedule = [];
  for (let y = 1; y <= a.holdYears; y++) {
    ebitda *= 1 + a.ebitdaGrowth;
    const interest = debt * a.interestRate;
    const capex = ebitda * a.capexPctOfEbitda;
    const preTax = ebitda - capex - interest;
    const tax = Math.max(0, preTax) * a.taxRate;
    const fcf = preTax - tax;
    const repayment = Math.max(0, Math.min(debt, fcf * a.cashSweepPct));
    debt -= repayment;
    schedule.push({
      year: y,
      ebitda: round1(ebitda),
      interest: round1(interest),
      capex: round1(capex),
      tax: round1(tax),
      fcf: round1(fcf),
      repayment: round1(repayment),
      debtEnd: round1(debt),
      leverage: round2(debt / ebitda),
    });
  }

  const exitEv = ebitda * a.exitMultiple;
  const exitEquity = exitEv - debt;
  const cashflows = [-equity0, ...Array(a.holdYears - 1).fill(0), exitEquity];
  const dealIrr = irr(cashflows);
  const dealMoic = moic(exitEquity, equity0);

  const sheets = [
    {
      name: 'Hypothèses',
      sectionTitle: `LBO — ${target}`,
      table: {
        headers: ['Hypothèse', 'Valeur'],
        rows: [
          ['EBITDA d\'entrée (m€)', a.entryEbitda],
          ['Multiple d\'entrée (VE/EBITDA)', `${a.entryMultiple}x`],
          ['Multiple de sortie', `${a.exitMultiple}x`],
          ['Durée de détention (années)', a.holdYears],
          ['% Dette à l\'entrée', formatPercent(a.debtPct)],
          ['Croissance EBITDA', formatPercent(a.ebitdaGrowth)],
          ['Taux d\'intérêt moyen', formatPercent(a.interestRate)],
          ['Cash sweep', formatPercent(a.cashSweepPct)],
          ['Taux d\'impôt', formatPercent(a.taxRate)],
          ['Frais de transaction', formatPercent(a.transactionFeesPct)],
        ],
      },
      columns: [{ width: 36 }, { width: 18 }],
    },
    {
      name: 'Sources & Emplois',
      sectionTitle: 'Sources & Emplois',
      table: {
        headers: ['Sources', 'm€', 'Emplois', 'm€'],
        rows: [
          ['Dette d\'acquisition', round1(debt0), 'Valeur d\'entreprise', round1(ev)],
          ['Fonds propres sponsor', round1(equity0), 'Frais de transaction', round1(fees)],
          ['Total sources', round1(debt0 + equity0), 'Total emplois', round1(ev + fees)],
        ],
        totalRowIndex: 2,
      },
      columns: [{ width: 28 }, { width: 14 }, { width: 28 }, { width: 14 }],
    },
    {
      name: 'Debt Schedule',
      sectionTitle: 'Échéancier de dette et flux',
      table: {
        headers: ['Année', 'EBITDA', 'Intérêts', 'Capex', 'Impôt', 'FCF', 'Remboursement', 'Dette fin', 'Levier'],
        rows: schedule.map((s) => [s.year, s.ebitda, s.interest, s.capex, s.tax, s.fcf, s.repayment, s.debtEnd, `${s.leverage}x`]),
      },
      columns: [{ width: 10 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 16 }, { width: 12 }, { width: 10 }],
      freezeHeader: true,
    },
    {
      name: 'Retours',
      sectionTitle: 'Analyse des retours sponsor',
      table: {
        headers: ['Métrique', 'Valeur'],
        rows: [
          ['Equity investi (m€)', round1(equity0)],
          ['VE de sortie (m€)', round1(exitEv)],
          ['Dette résiduelle (m€)', round1(debt)],
          ['Equity de sortie (m€)', round1(exitEquity)],
          ['MoM', dealMoic != null ? formatMultiple(dealMoic, { decimals: 2 }) : '—'],
          ['TRI', dealIrr != null ? formatPercent(dealIrr) : '—'],
        ],
      },
      columns: [{ width: 30 }, { width: 18 }],
    },
    {
      name: 'Sensibilités',
      sectionTitle: 'TRI selon multiple de sortie x croissance EBITDA',
      table: buildSensitivity(a, equity0, debt0),
      freezeHeader: true,
    },
  ];

  return buildWorkbook({ title: `LBO — ${target}`, sheets });
}

function round1(v) { return Math.round(v * 10) / 10; }
function round2(v) { return Math.round(v * 100) / 100; }

function buildSensitivity(a, equity0) {
  const exitMultiples = [-1, -0.5, 0, 0.5, 1].map((d) => a.exitMultiple + d);
  const growths = [-0.02, -0.01, 0, 0.01, 0.02].map((d) => a.ebitdaGrowth + d);
  const rows = exitMultiples.map((xm) => {
    const cells = growths.map((g) => {
      const r = simulateIrr({ ...a, exitMultiple: xm, ebitdaGrowth: g });
      return r == null ? '—' : formatPercent(r);
    });
    return [`${xm}x`, ...cells];
  });
  return {
    headers: ['Sortie \\ Croissance', ...growths.map((g) => formatPercent(g))],
    rows,
  };
}

function simulateIrr(a) {
  const ev = a.entryEbitda * a.entryMultiple;
  const debt0 = ev * a.debtPct;
  const equity0 = ev * (1 + a.transactionFeesPct) - debt0;
  let ebitda = a.entryEbitda;
  let debt = debt0;
  for (let y = 1; y <= a.holdYears; y++) {
    ebitda *= 1 + a.ebitdaGrowth;
    const interest = debt * a.interestRate;
    const capex = ebitda * a.capexPctOfEbitda;
    const preTax = ebitda - capex - interest;
    const tax = Math.max(0, preTax) * a.taxRate;
    const fcf = preTax - tax;
    debt -= Math.max(0, Math.min(debt, fcf * a.cashSweepPct));
  }
  const exitEquity = ebitda * a.exitMultiple - debt;
  return irr([-equity0, ...Array(a.holdYears - 1).fill(0), exitEquity]);
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
