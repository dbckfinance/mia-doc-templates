// generators/ma/teaser.js
// Deal teaser anonymisé (pptx, 2-3 slides)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildDeck } from '../../shared/pptx-deck.js';
import { disclaimerFor } from '../../shared/disclaimer-library.js';
import { formatCurrency, formatPercent } from '../../shared/financial-helpers.js';

export const metadata = {
  id: 'teaser',
  name: 'Deal Teaser (anonymisé)',
  vertical: 'ma',
  outputType: 'pptx',
  estimatedPages: '2-3 slides',
  requiredInput: ['codeName'],
  optionalInput: ['sector', 'description', 'highlights', 'financials', 'years', 'transaction', 'contact'],
};

export async function generate(input = {}) {
  const codeName = input.codeName || 'Projet Confidentiel';
  const fin = input.financials || {};
  const years = input.years || [];
  const slides = [];

  // Opportunity overview
  slides.push({
    type: 'content',
    title: "Opportunité d'investissement",
    paragraphs: input.description ? [input.description] : [],
    bullets: input.highlights || [
      'Position de marché solide sur un segment en croissance',
      'Base de clients récurrente et diversifiée',
      'Équipe de management expérimentée',
      'Multiples leviers de création de valeur identifiés',
    ],
  });

  // Financial profile
  if (fin.revenue?.length && years.length) {
    slides.push({
      type: 'content',
      title: 'Profil financier',
      twoCol: {
        left: {
          chart: {
            kind: 'bar',
            title: "Chiffre d'affaires (m€)",
            series: [{ name: 'CA', labels: years.map(String), values: fin.revenue }],
          },
        },
        right: {
          bullets: [
            fin.revenue ? `CA ${years[years.length - 1]}: ${formatCurrency(fin.revenue[fin.revenue.length - 1] * 1e6)}` : null,
            fin.ebitda ? `EBITDA: ${formatCurrency(fin.ebitda[fin.ebitda.length - 1] * 1e6)}` : null,
            fin.ebitda && fin.revenue ? `Marge EBITDA: ${formatPercent(fin.ebitda[fin.ebitda.length - 1] / fin.revenue[fin.revenue.length - 1])}` : null,
            ...(input.financialHighlights || []),
          ].filter(Boolean),
        },
      },
    });
  }

  // Transaction
  slides.push({
    type: 'facts',
    title: 'Processus envisagé',
    facts: [
      ['Nature de la transaction', input.transaction?.type || 'Cession majoritaire'],
      ['Calendrier indicatif', input.transaction?.timeline || 'À déterminer'],
      ['Processus', input.transaction?.process || 'Processus compétitif restreint'],
      ['Contact', input.contact || 'M&IA — équipe Transaction'],
    ],
  });

  return buildDeck({
    deckTitle: codeName,
    deckSubtitle: input.sector ? `Opportunité — ${input.sector}` : "Opportunité d'investissement",
    date: input.date,
    lang: 'fr',
    slides,
    disclaimer: disclaimerFor('ma', metadata.id),
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
