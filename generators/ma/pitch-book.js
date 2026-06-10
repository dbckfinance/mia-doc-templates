// generators/ma/pitch-book.js
// Pitch book / credentials deck (pptx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildDeck } from '../../shared/pptx-deck.js';
import { disclaimerFor } from '../../shared/disclaimer-library.js';

export const metadata = {
  id: 'pitch-book',
  name: 'Pitch Book / Credentials',
  vertical: 'ma',
  outputType: 'pptx',
  estimatedPages: '15-30 slides',
  requiredInput: ['client'],
  optionalInput: ['situation', 'opportunities', 'credentials', 'team', 'valuation', 'processOptions', 'nextSteps'],
};

export async function generate(input = {}) {
  const client = input.client || 'Client';
  const slides = [];

  slides.push({ type: 'section', title: 'Notre compréhension de votre situation' });
  slides.push({
    type: 'content',
    title: 'Contexte et enjeux',
    paragraphs: input.situation ? [input.situation] : [],
    bullets: input.issues || [],
  });

  if (input.opportunities?.length) {
    slides.push({ type: 'section', title: 'Opportunités stratégiques' });
    slides.push({ type: 'content', title: 'Options envisageables', bullets: input.opportunities });
  }

  if (input.valuation?.ranges?.length) {
    slides.push({ type: 'section', title: 'Éléments de valorisation' });
    slides.push({
      type: 'chart',
      title: 'Fourchette de valorisation indicative',
      chart: { kind: 'footballField', ranges: input.valuation.ranges, currency: input.valuation.currency || '€' },
    });
  }

  if (input.processOptions?.length) {
    slides.push({ type: 'section', title: 'Processus recommandé' });
    slides.push({
      type: 'table',
      title: 'Comparaison des options de processus',
      table: {
        headers: ['Option', 'Avantages', 'Inconvénients', 'Durée'],
        rows: input.processOptions.map((p) => [p.name, p.pros || '—', p.cons || '—', p.duration || '—']),
      },
    });
  }

  slides.push({ type: 'section', title: 'Pourquoi M&IA' });
  if (input.credentials?.length) {
    slides.push({
      type: 'table',
      title: 'Références récentes',
      table: {
        headers: ['Transaction', 'Secteur', 'Rôle', 'Année'],
        rows: input.credentials.map((c) => [c.deal, c.sector || '—', c.role || 'Conseil', c.year || '—']),
      },
    });
  }
  if (input.team?.length) {
    slides.push({
      type: 'table',
      title: 'Équipe dédiée',
      table: {
        headers: ['Nom', 'Rôle', 'Expérience'],
        rows: input.team.map((t) => [t.name, t.role, t.bio || '—']),
      },
    });
  }

  slides.push({
    type: 'content',
    title: 'Prochaines étapes',
    bullets: input.nextSteps || [
      'Validation du mandat et signature de la lettre de mission',
      'Préparation de la documentation marketing',
      'Lancement du processus',
    ],
  });

  return buildDeck({
    deckTitle: 'Pitch — Opportunité Stratégique',
    deckSubtitle: client,
    project: input.project,
    date: input.date,
    lang: 'fr',
    slides,
    disclaimer: disclaimerFor('ma', metadata.id),
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
