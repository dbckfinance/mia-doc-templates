// generators/ib/debt-offering-memo.js
// Bond / credit facility memo (pptx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildDeck } from '../../shared/pptx-deck.js';
import { disclaimerFor } from '../../shared/disclaimer-library.js';
import { creditMetrics } from '../../shared/financial-helpers.js';

export const metadata = {
  id: 'debt-offering-memo',
  name: 'Debt Offering Memo (Bond / Facility)',
  vertical: 'ib',
  outputType: 'pptx',
  estimatedPages: '12-20 slides',
  requiredInput: ['issuer'],
  optionalInput: ['terms', 'creditHighlights', 'financials', 'covenants', 'comparables', 'rating'],
};

export async function generate(input = {}) {
  const issuer = input.issuer || 'Émetteur';
  const terms = input.terms || {};
  const slides = [];

  slides.push({
    type: 'facts',
    title: 'Termes indicatifs',
    facts: [
      ['Émetteur', issuer],
      ['Instrument', terms.instrument || 'Obligations senior'],
      ['Montant', terms.amount || '—'],
      ['Maturité', terms.maturity || '—'],
      ['Coupon indicatif', terms.coupon || '—'],
      ['Notation', input.rating || 'Non notée'],
      ['Sûretés', terms.security || '—'],
      ['Use of proceeds', terms.useOfProceeds || 'Refinancement'],
    ],
  });

  if (input.creditHighlights?.length) {
    slides.push({ type: 'section', title: 'Credit Story' });
    slides.push({ type: 'content', title: 'Points forts du crédit', bullets: input.creditHighlights });
  }

  const fin = input.financials || {};
  const metrics = creditMetrics(fin);
  if (metrics.length) {
    slides.push({
      type: 'table',
      title: 'Métriques de crédit',
      table: { headers: ['Métrique', 'Valeur'], rows: metrics },
    });
  }

  if (input.covenants?.length) {
    slides.push({
      type: 'table',
      title: 'Covenants',
      table: {
        headers: ['Covenant', 'Seuil', 'Niveau actuel', 'Headroom'],
        rows: input.covenants.map((c) => [c.name, c.threshold || '—', c.current || '—', c.headroom || '—']),
      },
    });
  }

  if (input.comparables?.length) {
    slides.push({
      type: 'table',
      title: 'Émissions comparables',
      table: {
        headers: ['Émetteur', 'Instrument', 'Montant', 'Maturité', 'Spread / Coupon'],
        rows: input.comparables.map((c) => [c.name, c.instrument || '—', c.amount || '—', c.maturity || '—', c.pricing || '—']),
      },
    });
  }

  return buildDeck({
    deckTitle: 'Debt Offering Memorandum',
    deckSubtitle: issuer,
    project: input.project,
    date: input.date,
    lang: 'fr',
    slides,
    disclaimer: disclaimerFor('ib', metadata.id),
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
