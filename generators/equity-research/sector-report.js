// generators/equity-research/sector-report.js
// Rapport thématique sectoriel (pptx + docx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildDeck } from '../../shared/pptx-deck.js';
import { buildReport } from '../../shared/docx-report.js';

export const metadata = {
  id: 'sector-report',
  name: 'Rapport Sectoriel Thématique',
  vertical: 'equity-research',
  outputType: 'multi',
  outputs: ['pptx', 'docx'],
  estimatedPages: '8 slides + 10 pages',
  requiredInput: ['sector'],
  optionalInput: ['theme', 'keyTrends', 'coverage', 'topPicks', 'marketData', 'analyst'],
};

export async function generate(input = {}) {
  const sector = input.sector || 'Secteur';
  const theme = input.theme || `Perspectives ${new Date().getFullYear()}`;
  const trends = input.keyTrends || [
    'Consolidation en cours — multiplication des opérations de M&A',
    'Pression réglementaire croissante redessinant les modèles économiques',
    'Digitalisation accélérée créant des écarts de performance durables',
    'Normalisation des chaînes d\'approvisionnement',
  ];
  const coverage = input.coverage || [
    { name: 'Valeur A', rating: 'Achat', target: '45 €', upside: '+22%' },
    { name: 'Valeur B', rating: 'Neutre', target: '78 €', upside: '+4%' },
    { name: 'Valeur C', rating: 'Achat', target: '31 €', upside: '+18%' },
    { name: 'Valeur D', rating: 'Vente', target: '12 €', upside: '-15%' },
  ];
  const topPicks = input.topPicks || coverage.filter((c) => c.rating === 'Achat').map((c) => c.name);

  const pptx = await buildDeck({
    title: `${sector} — ${theme}`,
    subtitle: 'Rapport sectoriel',
    confidential: false,
    slides: [
      { type: 'section', title: 'Vue d\'ensemble du secteur' },
      { type: 'content', title: 'Tendances structurantes', bullets: trends },
      {
        type: 'table',
        title: 'Univers de couverture',
        headers: ['Valeur', 'Recommandation', 'Objectif', 'Potentiel'],
        rows: coverage.map((c) => [c.name, c.rating, c.target, c.upside]),
      },
      { type: 'content', title: 'Nos valeurs préférées', bullets: topPicks.map((p) => `${p} — voir détail dans le rapport complet`) },
    ],
  });

  const docx = await buildReport({
    docTitle: `${sector} — ${theme}`,
    docSubtitle: 'Rapport sectoriel thématique',
    vertical: 'equity-research',
    docId: 'sector-report',
    confidential: false,
    toc: true,
    sections: [
      {
        heading: 'Synthèse',
        blocks: [
          { type: 'p', text: input.summary || `Ce rapport analyse les dynamiques du secteur ${sector} et leurs implications pour les valeurs de notre univers de couverture.` },
          { type: 'bullets', items: trends },
        ],
      },
      {
        heading: 'Univers de couverture et recommandations',
        blocks: [
          {
            type: 'table',
            headers: ['Valeur', 'Recommandation', 'Objectif de cours', 'Potentiel'],
            rows: coverage.map((c) => [c.name, c.rating, c.target, c.upside]),
          },
        ],
      },
      {
        heading: 'Analyse des tendances',
        blocks: trends.map((t, i) => ({ type: 'p', text: `${i + 1}. ${t} — développement détaillé de la tendance, de ses moteurs et de ses gagnants/perdants dans notre univers.` })),
      },
      {
        heading: 'Valeurs préférées',
        blocks: [
          { type: 'bullets', items: topPicks },
        ],
      },
    ],
  });

  return [
    { fileName: 'sector-report.pptx', buffer: pptx, ext: 'pptx' },
    { fileName: 'sector-report.docx', buffer: docx, ext: 'docx' },
  ];
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
