// generators/ma/management-presentation.js
// Management presentation (pptx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildDeck } from '../../shared/pptx-deck.js';
import { disclaimerFor } from '../../shared/disclaimer-library.js';

export const metadata = {
  id: 'management-presentation',
  name: 'Management Presentation',
  vertical: 'ma',
  outputType: 'pptx',
  estimatedPages: '20-40 slides',
  requiredInput: ['company'],
  optionalInput: ['agenda', 'sections', 'financials', 'years', 'management', 'qa'],
};

export async function generate(input = {}) {
  const company = input.company || 'Société';
  const slides = [];

  slides.push({
    type: 'content',
    title: 'Agenda',
    bullets: input.agenda || [
      'Présentation de l\'équipe',
      'Activité et positionnement',
      'Performance financière',
      'Plan de développement',
      'Questions / réponses',
    ],
  });

  if (input.management?.length) {
    slides.push({ type: 'section', title: 'Équipe dirigeante' });
    slides.push({
      type: 'table',
      title: 'Intervenants du jour',
      table: {
        headers: ['Nom', 'Fonction', 'Ancienneté', 'Parcours'],
        rows: input.management.map((m) => [m.name, m.role, m.tenure || '—', m.bio || '—']),
      },
    });
  }

  for (const s of input.sections || []) {
    if (s.isSection) slides.push({ type: 'section', title: s.title });
    else {
      slides.push({
        type: 'content',
        title: s.title,
        paragraphs: s.paragraphs,
        bullets: s.bullets,
        table: s.table,
        chart: s.chart,
      });
    }
  }

  const fin = input.financials || {};
  const years = input.years || [];
  if (fin.revenue?.length && years.length) {
    slides.push({ type: 'section', title: 'Performance financière' });
    slides.push({
      type: 'chart',
      title: 'Trajectoire financière (m€)',
      chart: {
        kind: 'bar',
        series: [
          { name: 'CA', labels: years.map(String), values: fin.revenue },
          ...(fin.ebitda ? [{ name: 'EBITDA', labels: years.map(String), values: fin.ebitda }] : []),
        ],
      },
    });
  }

  if (input.qa?.length) {
    slides.push({ type: 'section', title: 'Questions / Réponses' });
    slides.push({ type: 'content', title: 'Thèmes anticipés', bullets: input.qa });
  }

  return buildDeck({
    deckTitle: 'Management Presentation',
    deckSubtitle: company,
    project: input.project,
    date: input.date,
    lang: 'fr',
    slides,
    disclaimer: disclaimerFor('ma', metadata.id),
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
