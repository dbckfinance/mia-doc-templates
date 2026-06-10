// generators/hedge-fund/event-driven-memo.js
// Mémo event-driven / situations spéciales (docx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildReport } from '../../shared/docx-report.js';

export const metadata = {
  id: 'event-driven-memo',
  name: 'Mémo Event-Driven / Situations Spéciales',
  vertical: 'hedge-fund',
  outputType: 'docx',
  estimatedPages: '5-10',
  requiredInput: ['situation'],
  optionalInput: ['eventType', 'company', 'spread', 'probability', 'timeline', 'scenarios', 'risks', 'positionStructure'],
};

export async function generate(input = {}) {
  const situation = input.situation || 'Situation spéciale';
  const company = input.company || situation;
  const eventType = input.eventType || 'Merger arbitrage';
  const scenarios = input.scenarios || [
    { name: 'Closing du deal aux conditions annoncées', probability: 0.75, payoff: '+6,5%' },
    { name: 'Sur-enchère d\'un tiers', probability: 0.10, payoff: '+14%' },
    { name: 'Échec du deal (antitrust)', probability: 0.15, payoff: '-18%' },
  ];

  return buildReport({
    docTitle: `Mémo event-driven — ${company}`,
    docSubtitle: eventType,
    vertical: 'hedge-fund',
    docId: 'event-driven-memo',
    confidential: true,
    sections: [
      {
        heading: 'Résumé de la situation',
        blocks: [
          { type: 'facts', facts: [
            { label: 'Type d\'événement', value: eventType },
            { label: 'Spread actuel', value: input.spread || 'n.d.' },
            { label: 'Probabilité de succès estimée', value: input.probability || '75%' },
            { label: 'Horizon', value: input.timeline || '3-9 mois' },
          ] },
          { type: 'p', text: input.summary || `Analyse de l'opportunité ${eventType.toLowerCase()} sur ${company}, incluant l'arbre de scénarios, l'espérance de gain et la structuration de la position.` },
        ],
      },
      {
        heading: 'Description de l\'événement',
        blocks: [
          { type: 'p', text: input.eventDescription || 'Description de l\'opération : termes annoncés, parties prenantes, conditions suspensives (antitrust, vote des actionnaires, financement), calendrier réglementaire.' },
        ],
      },
      {
        heading: 'Arbre de scénarios',
        blocks: [
          {
            type: 'table',
            headers: ['Scénario', 'Probabilité', 'Payoff estimé'],
            rows: scenarios.map((s) => [s.name, typeof s.probability === 'number' ? `${Math.round(s.probability * 100)}%` : s.probability, s.payoff]),
          },
          { type: 'p', text: 'L\'espérance de gain pondérée par les probabilités justifie la mise en place de la position au niveau de spread actuel.' },
        ],
      },
      {
        heading: 'Structuration de la position',
        blocks: [
          { type: 'p', text: input.positionStructure || 'Long cible / short acquéreur selon la parité d\'échange. Couverture du risque de marché via futures. Taille initiale de 2% NAV avec montée possible à 4% après franchissement des jalons réglementaires.' },
        ],
      },
      {
        heading: 'Risques',
        blocks: [
          { type: 'bullets', items: input.risks || [
            'Risque réglementaire (antitrust, CFIUS, autorités locales)',
            'Risque de financement de l\'acquéreur',
            'Risque de marché en cas de rupture du deal (gap down)',
            'Risque de calendrier (allongement = baisse du TRI annualisé)',
          ] },
        ],
      },
    ],
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
