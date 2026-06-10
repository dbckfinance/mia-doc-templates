// generators/ma/deal-summary.js
// Post-deal summary / tombstone deck (pptx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildDeck } from '../../shared/pptx-deck.js';
import { disclaimerFor } from '../../shared/disclaimer-library.js';
import { formatCurrency, formatMultiple } from '../../shared/financial-helpers.js';

export const metadata = {
  id: 'deal-summary',
  name: 'Deal Summary / Tombstone',
  vertical: 'ma',
  outputType: 'pptx',
  estimatedPages: '3-6 slides',
  requiredInput: ['target'],
  optionalInput: ['buyer', 'seller', 'metrics', 'highlights', 'role', 'timeline'],
};

export async function generate(input = {}) {
  const target = input.target || 'Société';
  const slides = [];

  slides.push({
    type: 'facts',
    title: 'Synthèse de la transaction',
    facts: [
      ['Cible', target],
      ['Acquéreur', input.buyer || '—'],
      ['Cédant', input.seller || '—'],
      ['Valeur d\'entreprise', input.metrics?.ev ? formatCurrency(input.metrics.ev) : 'Non communiquée'],
      ['VE / EBITDA', input.metrics?.evEbitda ? formatMultiple(input.metrics.evEbitda) : '—'],
      ['Rôle de M&IA', input.role || 'Conseil financier exclusif du cédant'],
      ['Date de closing', input.closingDate || '—'],
    ],
  });

  if (input.highlights?.length) {
    slides.push({ type: 'content', title: 'Points clés de la transaction', bullets: input.highlights });
  }

  if (input.timeline?.length) {
    slides.push({
      type: 'table',
      title: 'Déroulé du processus',
      table: {
        headers: ['Étape', 'Date', 'Commentaire'],
        rows: input.timeline.map((t) => [t.step, t.date || '—', t.notes || '']),
      },
    });
  }

  if (input.lessons?.length) {
    slides.push({ type: 'content', title: 'Enseignements', bullets: input.lessons });
  }

  return buildDeck({
    deckTitle: `Transaction réalisée — ${target}`,
    deckSubtitle: input.buyer ? `Acquisition par ${input.buyer}` : 'Deal Summary',
    date: input.date,
    lang: 'fr',
    slides,
    disclaimer: disclaimerFor('ma', metadata.id),
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
