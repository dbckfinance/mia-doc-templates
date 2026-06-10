// generators/hedge-fund/risk-report.js
// Rapport de risque : VaR, drawdown, expositions (xlsx + pptx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildWorkbook } from '../../shared/xlsx-model.js';
import { buildDeck } from '../../shared/pptx-deck.js';
import { formatPercent } from '../../shared/financial-helpers.js';

export const metadata = {
  id: 'risk-report',
  name: 'Rapport de Risque',
  vertical: 'hedge-fund',
  outputType: 'multi',
  outputs: ['xlsx', 'pptx'],
  estimatedPages: '3 sheets + 3 slides',
  requiredInput: ['fund'],
  optionalInput: ['period', 'var', 'drawdown', 'exposuresBySector', 'stressTests', 'limits'],
};

export async function generate(input = {}) {
  const fund = input.fund || 'Fonds';
  const period = input.period || new Date().toLocaleDateString('fr-FR');
  const varData = input.var || { var95: 0.014, var99: 0.024, expectedShortfall: 0.031, horizon: '1 jour' };
  const dd = input.drawdown || { current: -0.021, max: -0.085, maxDate: '2025-08', recoveryDays: 47 };
  const sectors = input.exposuresBySector || [
    { sector: 'Technologie', long: 0.28, short: -0.12 },
    { sector: 'Santé', long: 0.18, short: -0.08 },
    { sector: 'Industrie', long: 0.15, short: -0.10 },
    { sector: 'Consommation', long: 0.12, short: -0.14 },
    { sector: 'Financières', long: 0.10, short: -0.05 },
  ];
  const stress = input.stressTests || [
    { scenario: 'Krach actions -20%', impact: -0.062 },
    { scenario: 'Taux +100 bps', impact: -0.018 },
    { scenario: 'Écartement spreads crédit', impact: -0.024 },
    { scenario: 'Crise de liquidité 2008-like', impact: -0.089 },
  ];
  const limits = input.limits || [
    { name: 'VaR 95% max', limit: '2,0%', current: formatPercent(varData.var95), status: varData.var95 <= 0.02 ? 'OK' : 'DÉPASSEMENT' },
    { name: 'Exposition brute max', limit: '200%', current: '145%', status: 'OK' },
    { name: 'Concentration single name max', limit: '10%', current: '8%', status: 'OK' },
  ];

  const xlsx = await buildWorkbook({
    title: `Risque — ${fund}`,
    sheets: [
      {
        name: 'VaR & Drawdown',
        sectionTitle: `Métriques de risque — ${period}`,
        table: {
          headers: ['Métrique', 'Valeur'],
          rows: [
            [`VaR 95% (${varData.horizon})`, formatPercent(varData.var95)],
            [`VaR 99% (${varData.horizon})`, formatPercent(varData.var99)],
            ['Expected Shortfall 97,5%', formatPercent(varData.expectedShortfall)],
            ['Drawdown actuel', formatPercent(dd.current)],
            ['Drawdown maximum', formatPercent(dd.max)],
            ['Date du drawdown max', dd.maxDate],
            ['Jours de récupération', dd.recoveryDays],
          ],
        },
        columns: [{ width: 32 }, { width: 16 }],
      },
      {
        name: 'Expositions',
        sectionTitle: 'Expositions par secteur',
        table: {
          headers: ['Secteur', 'Long', 'Short', 'Net'],
          rows: sectors.map((s) => [s.sector, formatPercent(s.long), formatPercent(s.short), formatPercent(s.long + s.short)]),
        },
        columns: [{ width: 22 }, { width: 12 }, { width: 12 }, { width: 12 }],
        freezeHeader: true,
      },
      {
        name: 'Stress tests',
        sectionTitle: 'Tests de résistance',
        table: {
          headers: ['Scénario', 'Impact NAV estimé'],
          rows: stress.map((s) => [s.scenario, formatPercent(s.impact)]),
        },
        columns: [{ width: 36 }, { width: 20 }],
      },
    ],
  });

  const pptx = await buildDeck({
    title: `Rapport de risque — ${fund}`,
    subtitle: period,
    confidential: true,
    slides: [
      {
        type: 'facts',
        title: 'Synthèse risque',
        facts: [
          { label: 'VaR 95%', value: formatPercent(varData.var95) },
          { label: 'VaR 99%', value: formatPercent(varData.var99) },
          { label: 'Drawdown actuel', value: formatPercent(dd.current) },
          { label: 'Drawdown max', value: formatPercent(dd.max) },
        ],
      },
      {
        type: 'chart',
        title: 'Expositions nettes par secteur',
        chart: {
          kind: 'bar',
          categories: sectors.map((s) => s.sector),
          series: [
            { name: 'Long', values: sectors.map((s) => Math.round(s.long * 100)) },
            { name: 'Short', values: sectors.map((s) => Math.round(s.short * 100)) },
          ],
          axisFormat: '0"%"',
        },
      },
      {
        type: 'table',
        title: 'Limites de risque',
        headers: ['Limite', 'Seuil', 'Niveau actuel', 'Statut'],
        rows: limits.map((l) => [l.name, l.limit, l.current, l.status]),
      },
    ],
  });

  return [
    { fileName: 'risk-report.xlsx', buffer: xlsx, ext: 'xlsx' },
    { fileName: 'risk-report.pptx', buffer: pptx, ext: 'pptx' },
  ];
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
