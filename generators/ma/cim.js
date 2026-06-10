// generators/ma/cim.js
// Confidential Information Memorandum (pptx, 30-60 slides)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildDeck } from '../../shared/pptx-deck.js';
import { disclaimerFor } from '../../shared/disclaimer-library.js';
import { formatPercent } from '../../shared/financial-helpers.js';

export const metadata = {
  id: 'cim',
  name: 'Confidential Information Memorandum',
  vertical: 'ma',
  outputType: 'pptx',
  estimatedPages: '30-60 slides',
  requiredInput: ['company'],
  optionalInput: ['sector', 'sections', 'financials', 'years', 'market', 'management', 'strategy', 'transaction'],
};

export async function generate(input = {}) {
  const company = input.company || 'Société';
  const fin = input.financials || {};
  const years = input.years || [];
  const slides = [];

  // 1. Executive summary
  slides.push({ type: 'section', title: 'Synthèse exécutive' });
  slides.push({
    type: 'content',
    title: "L'opportunité en bref",
    bullets: input.investmentHighlights || [
      'Leader sur son marché avec des barrières à l\'entrée fortes',
      'Croissance soutenue et rentabilité supérieure aux pairs',
      'Potentiel de consolidation et de développement international',
    ],
  });

  // 2. Company overview
  slides.push({ type: 'section', title: "Présentation de l'entreprise" });
  slides.push({
    type: 'facts',
    title: 'Carte d\'identité',
    facts: [
      ['Société', company],
      ['Secteur', input.sector || '—'],
      ['Fondation', input.founded || '—'],
      ['Siège', input.headquarters || '—'],
      ['Effectifs', input.headcount || '—'],
      ['Actionnariat', input.ownership || '—'],
    ],
  });
  if (input.businessDescription) {
    slides.push({ type: 'content', title: 'Activité', paragraphs: [input.businessDescription], bullets: input.businessLines || [] });
  }
  if (input.history?.length) {
    slides.push({ type: 'content', title: 'Historique et étapes clés', bullets: input.history });
  }

  // 3. Market
  if (input.market) {
    slides.push({ type: 'section', title: 'Marché et positionnement' });
    slides.push({
      type: 'content',
      title: 'Dynamique de marché',
      paragraphs: input.market.overview ? [input.market.overview] : [],
      bullets: input.market.drivers || [],
    });
    if (input.market.competitors?.length) {
      slides.push({
        type: 'table',
        title: 'Paysage concurrentiel',
        table: {
          headers: ['Acteur', 'Positionnement', 'Part de marché'],
          rows: input.market.competitors.map((c) => [c.name, c.positioning || '—', c.share || '—']),
        },
      });
    }
  }

  // 4. Strategy
  if (input.strategy?.length) {
    slides.push({ type: 'section', title: 'Stratégie et plan de développement' });
    slides.push({ type: 'content', title: 'Axes stratégiques', bullets: input.strategy });
  }

  // 5. Management
  if (input.management?.length) {
    slides.push({ type: 'section', title: 'Équipe de management' });
    slides.push({
      type: 'table',
      title: 'Équipe dirigeante',
      table: {
        headers: ['Nom', 'Fonction', 'Expérience'],
        rows: input.management.map((m) => [m.name, m.role, m.bio || '—']),
      },
    });
  }

  // 6. Financials
  slides.push({ type: 'section', title: 'Performance financière' });
  if (fin.revenue?.length && years.length) {
    slides.push({
      type: 'chart',
      title: "Évolution du chiffre d'affaires et de l'EBITDA (m€)",
      chart: {
        kind: 'bar',
        series: [
          { name: 'CA', labels: years.map(String), values: fin.revenue },
          ...(fin.ebitda ? [{ name: 'EBITDA', labels: years.map(String), values: fin.ebitda }] : []),
        ],
      },
    });
    const rows = [];
    if (fin.revenue) rows.push(["Chiffre d'affaires (m€)", ...fin.revenue]);
    if (fin.ebitda) rows.push(['EBITDA (m€)', ...fin.ebitda]);
    if (fin.ebitda && fin.revenue) rows.push(['Marge EBITDA', ...fin.ebitda.map((e, i) => formatPercent(e / fin.revenue[i]))]);
    if (fin.netIncome) rows.push(['Résultat net (m€)', ...fin.netIncome]);
    if (fin.capex) rows.push(['Capex (m€)', ...fin.capex]);
    if (fin.netDebt) rows.push(['Dette nette (m€)', ...fin.netDebt]);
    slides.push({
      type: 'table',
      title: 'Synthèse financière',
      table: { headers: ['Agrégat', ...years.map(String)], rows },
    });
  }
  if (input.businessPlan) {
    slides.push({
      type: 'content',
      title: 'Business plan',
      paragraphs: input.businessPlan.commentary ? [input.businessPlan.commentary] : [],
      bullets: input.businessPlan.assumptions || [],
    });
  }

  // 7. Transaction
  slides.push({ type: 'section', title: 'Transaction envisagée' });
  slides.push({
    type: 'facts',
    title: 'Cadre de la transaction',
    facts: [
      ['Nature', input.transaction?.type || 'Cession majoritaire'],
      ['Périmètre', input.transaction?.scope || '100% du capital'],
      ['Processus', input.transaction?.process || 'Processus compétitif'],
      ['Calendrier', input.transaction?.timeline || 'À déterminer'],
      ['Conseil', 'M&IA'],
    ],
  });

  // Custom additional sections
  for (const s of input.sections || []) {
    slides.push({ type: 'content', title: s.title, paragraphs: s.paragraphs, bullets: s.bullets, table: s.table });
  }

  return buildDeck({
    deckTitle: 'Confidential Information Memorandum',
    deckSubtitle: company,
    project: input.project,
    date: input.date,
    lang: 'fr',
    slides,
    disclaimer: disclaimerFor('ma', metadata.id),
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
