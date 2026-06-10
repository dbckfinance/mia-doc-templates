// generators/ib/equity-offering-memo.js
// IPO / secondary offering memo (pptx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildDeck } from '../../shared/pptx-deck.js';
import { disclaimerFor } from '../../shared/disclaimer-library.js';

export const metadata = {
  id: 'equity-offering-memo',
  name: 'Equity Offering Memo (IPO / Secondary)',
  vertical: 'ib',
  outputType: 'pptx',
  estimatedPages: '15-25 slides',
  requiredInput: ['issuer'],
  optionalInput: ['offering', 'useOfProceeds', 'positioning', 'comparables', 'timeline', 'marketConditions'],
};

export async function generate(input = {}) {
  const issuer = input.issuer || 'Émetteur';
  const offering = input.offering || {};
  const slides = [];

  slides.push({
    type: 'facts',
    title: "Termes indicatifs de l'opération",
    facts: [
      ['Émetteur', issuer],
      ['Nature', offering.type || 'Introduction en bourse (IPO)'],
      ['Taille indicative', offering.size || '—'],
      ['Structure', offering.structure || 'Primaire / secondaire à déterminer'],
      ['Place de cotation', offering.exchange || 'Euronext Paris'],
      ['Calendrier cible', offering.timing || '—'],
    ],
  });

  if (input.positioning?.length) {
    slides.push({ type: 'section', title: 'Equity Story' });
    slides.push({ type: 'content', title: "Piliers de l'equity story", bullets: input.positioning });
  }

  if (input.marketConditions) {
    slides.push({
      type: 'content',
      title: 'Conditions de marché',
      paragraphs: input.marketConditions.commentary ? [input.marketConditions.commentary] : [],
      bullets: input.marketConditions.points || [],
    });
  }

  if (input.comparables?.length) {
    slides.push({
      type: 'table',
      title: 'Comparables cotés',
      table: {
        headers: ['Société', 'Capitalisation (m€)', 'VE/EBITDA', 'P/E', 'Performance YTD'],
        rows: input.comparables.map((c) => [c.name, c.marketCap ?? '—', c.evEbitda ? `${c.evEbitda}x` : '—', c.pe ? `${c.pe}x` : '—', c.ytd || '—']),
      },
    });
  }

  if (input.useOfProceeds?.length) {
    slides.push({
      type: 'chart',
      title: 'Utilisation des fonds',
      chart: {
        kind: 'pie',
        labels: input.useOfProceeds.map((u) => u.label),
        values: input.useOfProceeds.map((u) => u.value),
      },
    });
  }

  if (input.timeline?.length) {
    slides.push({
      type: 'table',
      title: 'Calendrier indicatif',
      table: {
        headers: ['Phase', 'Période', 'Jalons'],
        rows: input.timeline.map((t) => [t.phase, t.period || '—', t.milestones || '—']),
      },
    });
  }

  slides.push({
    type: 'content',
    title: 'Prochaines étapes',
    bullets: input.nextSteps || [
      'Désignation des banques du syndicat',
      'Kick-off et due diligence',
      'Préparation du prospectus',
      'Pilot fishing / early-look meetings',
    ],
  });

  return buildDeck({
    deckTitle: 'Equity Offering Memorandum',
    deckSubtitle: issuer,
    project: input.project,
    date: input.date,
    lang: 'fr',
    slides,
    disclaimer: disclaimerFor('ib', metadata.id),
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
