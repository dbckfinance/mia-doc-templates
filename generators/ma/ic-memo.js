// generators/ma/ic-memo.js
// Investment Committee memo (docx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildReport } from '../../shared/docx-report.js';
import { disclaimerFor } from '../../shared/disclaimer-library.js';
import { formatCurrency, formatMultiple, formatPercent } from '../../shared/financial-helpers.js';

export const metadata = {
  id: 'ic-memo',
  name: 'Investment Committee Memo',
  vertical: 'ma',
  outputType: 'docx',
  estimatedPages: '8-15',
  requiredInput: ['target'],
  optionalInput: ['recommendation', 'thesis', 'valuation', 'returns', 'risks', 'dd', 'financials', 'years'],
};

export async function generate(input = {}) {
  const target = input.target || 'Cible';
  const sections = [];

  sections.push({
    heading: 'Recommandation',
    paragraphs: [
      input.recommendation || `Il est proposé au Comité d'Investissement d'approuver la poursuite de l'opération relative à ${target} selon les termes décrits dans le présent mémorandum.`,
    ],
    facts: [
      ['Cible', target],
      ['Opération', input.dealType || 'Acquisition'],
      ['Valeur d\'entreprise', input.valuation?.ev ? formatCurrency(input.valuation.ev) : '—'],
      ['Multiple implicite', input.valuation?.evEbitda ? formatMultiple(input.valuation.evEbitda) : '—'],
      ['TRI cible', input.returns?.irr != null ? formatPercent(input.returns.irr) : '—'],
      ['Multiple cible (MoM)', input.returns?.moic != null ? formatMultiple(input.returns.moic) : '—'],
    ],
    pageBreakAfter: true,
  });

  sections.push({
    heading: "Thèse d'investissement",
    bullets: input.thesis || [
      'Position concurrentielle défendable',
      'Dynamique de marché favorable',
      'Leviers de création de valeur identifiés',
      'Voies de sortie multiples',
    ],
  });

  if (input.financials && input.years) {
    const fin = input.financials;
    const rows = [];
    if (fin.revenue) rows.push(["Chiffre d'affaires (m€)", ...fin.revenue]);
    if (fin.ebitda) rows.push(['EBITDA (m€)', ...fin.ebitda]);
    if (fin.ebitda && fin.revenue) rows.push(['Marge EBITDA', ...fin.ebitda.map((e, i) => formatPercent(e / fin.revenue[i]))]);
    sections.push({
      heading: 'Synthèse financière',
      table: { headers: ['Agrégat', ...input.years.map(String)], rows },
    });
  }

  if (input.valuation) {
    sections.push({
      heading: 'Valorisation',
      paragraphs: input.valuation.commentary ? [input.valuation.commentary] : [],
      table: input.valuation.methods?.length ? {
        headers: ['Méthode', 'Fourchette basse', 'Fourchette haute', 'Retenu'],
        rows: input.valuation.methods.map((m) => [m.name, m.low ?? '—', m.high ?? '—', m.selected ?? '—']),
      } : undefined,
    });
  }

  if (input.returns?.scenarios?.length) {
    sections.push({
      heading: 'Analyse des retours',
      table: {
        headers: ['Scénario', 'TRI', 'MoM', 'Hypothèses clés'],
        rows: input.returns.scenarios.map((s) => [
          s.name,
          s.irr != null ? formatPercent(s.irr) : '—',
          s.moic != null ? formatMultiple(s.moic) : '—',
          s.assumptions || '—',
        ]),
      },
    });
  }

  if (input.dd) {
    sections.push({
      heading: 'Conclusions des due diligences',
      bullets: input.dd.findings || [],
      paragraphs: input.dd.commentary ? [input.dd.commentary] : [],
    });
  }

  sections.push({
    heading: 'Risques et facteurs de mitigation',
    table: input.risks?.length ? {
      headers: ['Risque', 'Impact', 'Probabilité', 'Mitigation'],
      rows: input.risks.map((r) => [
        r.risk || r,
        r.impact || '—',
        r.likelihood || '—',
        r.mitigation || '—',
      ]),
    } : undefined,
    paragraphs: input.risks?.length ? [] : ['Analyse des risques à compléter.'],
  });

  if (input.nextSteps?.length) {
    sections.push({ heading: 'Prochaines étapes', bullets: input.nextSteps });
  }

  return buildReport({
    docTitle: 'Investment Committee Memorandum',
    docSubtitle: target,
    project: input.project,
    date: input.date,
    author: input.author,
    lang: 'fr',
    sections,
    disclaimer: disclaimerFor('ma', metadata.id),
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
