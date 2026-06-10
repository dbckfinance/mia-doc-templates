// generators/ma/note-synthese.js
// Note de synthèse financière (docx, 10-20 pages)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildReport } from '../../shared/docx-report.js';
import { disclaimerFor } from '../../shared/disclaimer-library.js';
import { buildVarianceTable, formatCurrency, formatPercent, cagr } from '../../shared/financial-helpers.js';

export const metadata = {
  id: 'note-synthese',
  name: 'Note de Synthèse Financière',
  vertical: 'ma',
  outputType: 'docx',
  estimatedPages: '10-20',
  requiredInput: ['company'],
  optionalInput: ['sector', 'financials', 'years', 'kbContent', 'comps', 'ratios', 'recommendations', 'risks', 'highlights'],
};

export async function generate(input = {}) {
  const company = input.company || 'Société cible';
  const years = input.years || [];
  const fin = input.financials || {};
  const sections = [];

  // Executive summary
  sections.push({
    heading: 'Synthèse exécutive',
    paragraphs: input.executiveSummary
      ? [input.executiveSummary]
      : [`La présente note de synthèse porte sur ${company}${input.sector ? `, acteur du secteur ${input.sector}` : ''}. Elle présente une analyse de la situation financière, de la performance opérationnelle et des principaux enjeux identifiés.`],
    bullets: input.highlights || [],
    pageBreakAfter: true,
  });

  // Company overview
  sections.push({
    heading: "Présentation de l'entreprise",
    paragraphs: input.companyOverview ? [input.companyOverview] : [],
    facts: [
      ['Société', company],
      ['Secteur', input.sector || '—'],
      ['Chiffre d\'affaires', fin.revenue?.length ? formatCurrency(fin.revenue[fin.revenue.length - 1] * 1e6) : '—'],
      ['EBITDA', fin.ebitda?.length ? formatCurrency(fin.ebitda[fin.ebitda.length - 1] * 1e6) : '—'],
      ['Effectifs', input.headcount || '—'],
      ['Actionnariat', input.ownership || '—'],
    ],
  });

  // Financial analysis
  const finRows = [];
  if (fin.revenue) finRows.push({ label: "Chiffre d'affaires (m€)", values: fin.revenue });
  if (fin.ebitda) finRows.push({ label: 'EBITDA (m€)', values: fin.ebitda });
  if (fin.ebitda && fin.revenue) {
    finRows.push({ label: 'Marge EBITDA', values: fin.ebitda.map((e, i) => e / fin.revenue[i]), isPercent: true });
  }
  if (fin.netIncome) finRows.push({ label: 'Résultat net (m€)', values: fin.netIncome });
  if (fin.netDebt) finRows.push({ label: 'Dette nette (m€)', values: fin.netDebt });
  if (fin.capex) finRows.push({ label: 'Capex (m€)', values: fin.capex });

  const finSection = {
    heading: 'Analyse financière',
    paragraphs: input.financialCommentary ? [input.financialCommentary] : [],
  };
  if (finRows.length && years.length) {
    finSection.table = { caption: 'Agrégats financiers historiques', ...buildVarianceTable(years, finRows) };
    const revCagr = fin.revenue?.length > 1 ? cagr(fin.revenue[0], fin.revenue[fin.revenue.length - 1], fin.revenue.length - 1) : null;
    if (revCagr != null) {
      finSection.paragraphs = [
        ...(finSection.paragraphs || []),
        `Le chiffre d'affaires affiche un TCAM de ${formatPercent(revCagr)} sur la période analysée.`,
      ];
    }
  }
  sections.push(finSection);

  // Ratios
  if (input.ratios?.length) {
    sections.push({
      heading: 'Ratios clés',
      table: {
        headers: ['Ratio', 'Valeur', 'Référence secteur', 'Lecture'],
        rows: input.ratios.map((r) => [r.name, r.value, r.benchmark || '—', r.comment || '']),
      },
    });
  }

  // Peer comparison
  if (input.comps?.length) {
    sections.push({
      heading: 'Comparaison sectorielle',
      table: {
        headers: ['Société', 'CA (m€)', 'Marge EBITDA', 'VE/EBITDA', 'VE/CA'],
        rows: input.comps.map((c) => [
          c.name,
          c.revenue ?? '—',
          c.ebitdaMargin != null ? formatPercent(c.ebitdaMargin) : '—',
          c.evEbitda != null ? `${c.evEbitda}x` : '—',
          c.evRevenue != null ? `${c.evRevenue}x` : '—',
        ]),
      },
    });
  }

  // KB content (extra analysis from M&IA knowledge base)
  if (input.kbContent) {
    sections.push({ heading: 'Éléments complémentaires', paragraphs: [input.kbContent] });
  }

  // Risks
  if (input.risks?.length) {
    sections.push({ heading: "Points d'attention et risques", bullets: input.risks });
  }

  // Recommendations
  if (input.recommendations?.length) {
    sections.push({ heading: 'Recommandations', bullets: input.recommendations });
  }

  return buildReport({
    docTitle: 'Note de Synthèse Financière',
    docSubtitle: company,
    project: input.project,
    date: input.date,
    author: input.author,
    lang: 'fr',
    sections,
    disclaimer: disclaimerFor('ma', metadata.id),
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
