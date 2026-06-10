// generators/common/client-presentation.js
// Présentation client générique (pptx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildDeck } from '../../shared/pptx-deck.js';

export const metadata = {
  id: 'client-presentation',
  name: 'Présentation Client',
  vertical: 'common',
  outputType: 'pptx',
  estimatedPages: 'variable',
  requiredInput: ['title'],
  optionalInput: ['subtitle', 'sections', 'confidential'],
};

export async function generate(input = {}) {
  const title = input.title || 'Présentation';

  // Si l'appelant fournit des sections structurées, on les rend telles quelles.
  const slides = [];
  const sections = input.sections || [
    {
      title: 'Introduction',
      slides: [
        { type: 'content', title: 'Objectifs de la réunion', bullets: ['Point d\'avancement', 'Décisions attendues', 'Prochaines étapes'] },
      ],
    },
    {
      title: 'Analyse',
      slides: [
        { type: 'content', title: 'Principaux constats', bullets: ['Constat 1', 'Constat 2', 'Constat 3'] },
      ],
    },
    {
      title: 'Recommandations',
      slides: [
        { type: 'content', title: 'Plan d\'action proposé', bullets: ['Action court terme', 'Action moyen terme', 'Jalon de décision'] },
      ],
    },
  ];

  for (const section of sections) {
    if (section.title) slides.push({ type: 'section', title: section.title });
    for (const slide of section.slides || []) slides.push(slide);
  }

  return buildDeck({
    title,
    subtitle: input.subtitle || '',
    confidential: input.confidential !== false,
    slides,
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
