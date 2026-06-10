// generators/ib/accretion-dilution.js
// Analyse de relution / dilution (xlsx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildWorkbook } from '../../shared/xlsx-model.js';
import { formatPercent } from '../../shared/financial-helpers.js';

export const metadata = {
  id: 'accretion-dilution',
  name: 'Analyse Relution / Dilution',
  vertical: 'ib',
  outputType: 'xlsx',
  estimatedPages: '3 sheets',
  requiredInput: ['acquirer', 'target'],
  optionalInput: ['acquirerNetIncome', 'acquirerShares', 'targetNetIncome', 'purchasePrice', 'pctCash', 'pctStock', 'costOfDebt', 'taxRate', 'synergies', 'acquirerSharePrice'],
};

export async function generate(input = {}) {
  const acquirer = input.acquirer || 'Acquéreur';
  const target = input.target || 'Cible';
  const a = {
    acquirerNetIncome: input.acquirerNetIncome ?? 200,
    acquirerShares: input.acquirerShares ?? 100,
    acquirerSharePrice: input.acquirerSharePrice ?? 35,
    targetNetIncome: input.targetNetIncome ?? 45,
    purchasePrice: input.purchasePrice ?? 600,
    pctCash: input.pctCash ?? 0.5,
    pctStock: input.pctStock ?? 0.5,
    costOfDebt: input.costOfDebt ?? 0.05,
    taxRate: input.taxRate ?? 0.25,
    synergies: input.synergies ?? 20,
  };

  const cashPortion = a.purchasePrice * a.pctCash;
  const stockPortion = a.purchasePrice * a.pctStock;
  const newShares = stockPortion / a.acquirerSharePrice;
  const afterTaxInterest = cashPortion * a.costOfDebt * (1 - a.taxRate);
  const afterTaxSynergies = a.synergies * (1 - a.taxRate);

  const epsStandalone = a.acquirerNetIncome / a.acquirerShares;
  const proFormaNI = a.acquirerNetIncome + a.targetNetIncome + afterTaxSynergies - afterTaxInterest;
  const proFormaShares = a.acquirerShares + newShares;
  const epsProForma = proFormaNI / proFormaShares;
  const accretion = epsProForma / epsStandalone - 1;

  // Sans synergies
  const epsNoSyn = (a.acquirerNetIncome + a.targetNetIncome - afterTaxInterest) / proFormaShares;
  const accretionNoSyn = epsNoSyn / epsStandalone - 1;

  return buildWorkbook({
    title: `Relution/Dilution — ${acquirer} / ${target}`,
    sheets: [
      {
        name: 'Hypothèses',
        sectionTitle: `${acquirer} acquiert ${target}`,
        table: {
          headers: ['Hypothèse', 'Valeur'],
          rows: [
            [`Résultat net ${acquirer} (m€)`, a.acquirerNetIncome],
            [`Nombre d'actions ${acquirer} (m)`, a.acquirerShares],
            [`Cours de l'action ${acquirer} (€)`, a.acquirerSharePrice],
            [`Résultat net ${target} (m€)`, a.targetNetIncome],
            ['Prix d\'acquisition (m€)', a.purchasePrice],
            ['% Cash', formatPercent(a.pctCash)],
            ['% Titres', formatPercent(a.pctStock)],
            ['Coût de la dette', formatPercent(a.costOfDebt)],
            ['Taux d\'impôt', formatPercent(a.taxRate)],
            ['Synergies avant impôt (m€)', a.synergies],
          ],
        },
        columns: [{ width: 38 }, { width: 16 }],
      },
      {
        name: 'Calcul',
        sectionTitle: 'Calcul pro forma',
        table: {
          headers: ['Ligne', 'm€ / m'],
          rows: [
            ['Financement cash (m€)', round1(cashPortion)],
            ['Financement titres (m€)', round1(stockPortion)],
            ['Nouvelles actions émises (m)', round2(newShares)],
            ['Intérêts après impôt (m€)', round1(afterTaxInterest)],
            ['Synergies après impôt (m€)', round1(afterTaxSynergies)],
            ['Résultat net pro forma (m€)', round1(proFormaNI)],
            ['Actions pro forma (m)', round2(proFormaShares)],
          ],
        },
        columns: [{ width: 38 }, { width: 16 }],
      },
      {
        name: 'Résultat',
        sectionTitle: 'Impact BPA',
        table: {
          headers: ['Métrique', 'Valeur'],
          rows: [
            ['BPA standalone (€)', round2(epsStandalone)],
            ['BPA pro forma avec synergies (€)', round2(epsProForma)],
            ['BPA pro forma sans synergies (€)', round2(epsNoSyn)],
            ['Relution / (dilution) avec synergies', formatPercent(accretion)],
            ['Relution / (dilution) sans synergies', formatPercent(accretionNoSyn)],
            ['Verdict', accretion >= 0 ? 'RELUTIF' : 'DILUTIF'],
          ],
        },
        columns: [{ width: 40 }, { width: 18 }],
      },
    ],
  });
}

function round1(v) { return Math.round(v * 10) / 10; }
function round2(v) { return Math.round(v * 100) / 100; }

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
