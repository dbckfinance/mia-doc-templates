// generators/hedge-fund/side-pocket-memo.js
// Mémo side pocket / position illiquide (docx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildReport } from '../../shared/docx-report.js';

export const metadata = {
  id: 'side-pocket-memo',
  name: 'Mémo Side Pocket / Position Illiquide',
  vertical: 'hedge-fund',
  outputType: 'docx',
  estimatedPages: '3-6',
  requiredInput: ['position'],
  optionalInput: ['fund', 'navImpact', 'valuationMethod', 'rationale', 'exitPlan', 'governance'],
};

export async function generate(input = {}) {
  const position = input.position || 'Position illiquide';
  const fund = input.fund || 'Fonds';

  return buildReport({
    docTitle: `Mémo side pocket — ${position}`,
    docSubtitle: fund,
    vertical: 'hedge-fund',
    docId: 'side-pocket-memo',
    confidential: true,
    sections: [
      {
        heading: 'Synthèse de la décision',
        blocks: [
          { type: 'kv', label: 'Position concernée', value: position },
          { type: 'kv', label: 'Part de la NAV', value: input.navImpact || 'n.d.' },
          { type: 'kv', label: 'Date d\'effet', value: input.effectiveDate || new Date().toLocaleDateString('fr-FR') },
          { type: 'p', text: input.rationale || `En raison de la perte de liquidité observable sur ${position} et de l'impossibilité d'en déterminer une juste valeur de marché fiable, la société de gestion a décidé de loger cette position dans un compartiment dédié (side pocket), conformément aux dispositions du règlement du fonds.` },
        ],
      },
      {
        heading: 'Contexte et justification',
        blocks: [
          { type: 'bullets', items: input.context || [
            'Suspension de la cotation / absence de transactions observables',
            'Horizon de réalisation incertain, incompatible avec la liquidité offerte aux porteurs',
            'Égalité de traitement des investisseurs entrants et sortants',
          ] },
        ],
      },
      {
        heading: 'Méthodologie de valorisation',
        blocks: [
          { type: 'p', text: input.valuationMethod || 'La position sera valorisée selon une approche multicritère (dernière transaction de référence, comparables, actualisation des flux attendus), revue trimestriellement par le comité de valorisation avec l\'appui d\'un expert indépendant le cas échéant.' },
        ],
      },
      {
        heading: 'Conséquences pour les porteurs',
        blocks: [
          { type: 'bullets', items: [
            'Les parts du compartiment principal restent souscriptibles et rachetables dans les conditions habituelles',
            'Les parts du side pocket ne sont ni souscriptibles ni rachetables ; elles seront remboursées au fur et à mesure des réalisations',
            'Aucune commission de gestion ne sera prélevée sur le side pocket au-delà des frais directs',
          ] },
        ],
      },
      {
        heading: 'Plan de sortie',
        blocks: [
          { type: 'p', text: input.exitPlan || 'La société de gestion poursuivra activement les options de monétisation : cession secondaire, restructuration, procédure judiciaire le cas échéant. Un point d\'avancement sera communiqué aux porteurs à chaque arrêté trimestriel.' },
        ],
      },
      {
        heading: 'Gouvernance',
        blocks: [
          { type: 'p', text: input.governance || 'La décision a été approuvée par le comité des risques et notifiée au dépositaire et au régulateur conformément à la réglementation applicable.' },
        ],
      },
    ],
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
