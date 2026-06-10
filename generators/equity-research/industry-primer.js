// generators/equity-research/industry-primer.js
// Primer sectoriel approfondi (pptx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildDeck } from '../../shared/pptx-deck.js';

export const metadata = {
  id: 'industry-primer',
  name: 'Industry Primer (Deep-Dive Sectoriel)',
  vertical: 'equity-research',
  outputType: 'pptx',
  estimatedPages: '15-30 slides',
  requiredInput: ['industry'],
  optionalInput: ['marketSize', 'segments', 'valueChain', 'competitiveLandscape', 'drivers', 'risks', 'keyPlayers', 'outlook'],
};

export async function generate(input = {}) {
  const industry = input.industry || 'Industrie';
  const segments = input.segments || [
    { name: 'Segment A', size: 45, growth: '+6%' },
    { name: 'Segment B', size: 30, growth: '+9%' },
    { name: 'Segment C', size: 25, growth: '+3%' },
  ];
  const players = input.keyPlayers || [
    { name: 'Leader 1', share: '22%', positioning: 'Leader global, intégré verticalement' },
    { name: 'Leader 2', share: '15%', positioning: 'Challenger, croissance par acquisitions' },
    { name: 'Leader 3', share: '11%', positioning: 'Spécialiste premium' },
  ];

  return buildDeck({
    title: `Primer — ${industry}`,
    subtitle: 'Deep-dive sectoriel',
    confidential: false,
    slides: [
      { type: 'section', title: '1. Vue d\'ensemble du marché' },
      {
        type: 'facts',
        title: 'Taille et croissance',
        facts: [
          { label: 'Taille de marché', value: input.marketSize || 'n.d.' },
          ...segments.map((s) => ({ label: s.name, value: `${s.size}% du marché, ${s.growth}/an` })),
        ],
      },
      {
        type: 'chart',
        title: 'Répartition par segment',
        chart: {
          kind: 'pie',
          labels: segments.map((s) => s.name),
          values: segments.map((s) => s.size),
        },
      },
      { type: 'section', title: '2. Chaîne de valeur' },
      {
        type: 'content',
        title: 'Chaîne de valeur',
        bullets: input.valueChain || [
          'Amont : fournisseurs de matières premières et composants',
          'Cœur : conception, production, assemblage',
          'Aval : distribution, services, après-vente',
          'Capture de valeur concentrée sur les segments à forte intensité technologique',
        ],
      },
      { type: 'section', title: '3. Paysage concurrentiel' },
      {
        type: 'table',
        title: 'Acteurs clés',
        headers: ['Acteur', 'Part de marché', 'Positionnement'],
        rows: players.map((p) => [p.name, p.share, p.positioning]),
      },
      { type: 'section', title: '4. Moteurs et risques' },
      {
        type: 'content',
        title: 'Moteurs de croissance',
        bullets: input.drivers || [
          'Croissance structurelle de la demande sous-jacente',
          'Innovation produit et montée en gamme',
          'Consolidation créant des économies d\'échelle',
        ],
      },
      {
        type: 'content',
        title: 'Risques sectoriels',
        bullets: input.risks || [
          'Cyclicité de la demande',
          'Pression réglementaire',
          'Disruption technologique',
        ],
      },
      {
        type: 'content',
        title: 'Perspectives',
        bullets: input.outlook || [
          'Croissance attendue supérieure au PIB sur le cycle',
          'Polarisation croissante entre leaders et acteurs sous-critiques',
          'Le pricing power restera le principal discriminant de performance',
        ],
      },
    ],
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
