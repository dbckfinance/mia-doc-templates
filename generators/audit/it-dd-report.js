// generators/audit/it-dd-report.js
// Rapport de due diligence IT (docx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildReport } from '../../shared/docx-report.js';

export const metadata = {
  id: 'it-dd-report',
  name: 'Due Diligence IT',
  vertical: 'audit',
  outputType: 'docx',
  estimatedPages: '10-15',
  requiredInput: ['company'],
  optionalInput: ['systems', 'findings', 'cyberPosture', 'capexNeeds'],
};

const DEFAULT_SYSTEMS = [
  { name: 'ERP', product: 'SAP S/4HANA', status: 'Moderne', comment: 'Migration achevée en 2024' },
  { name: 'CRM', product: 'Salesforce', status: 'Moderne', comment: 'Bien adopté par les équipes commerciales' },
  { name: 'Paie / RH', product: 'Solution legacy', status: 'Obsolète', comment: 'Fin de support éditeur dans 18 mois' },
  { name: 'Infrastructure', product: 'Hybride cloud / on-premise', status: 'Transition', comment: 'Migration cloud planifiée' },
];

const DEFAULT_FINDINGS = [
  { title: 'Dette technique sur le SI RH', severity: 'Élevé', detail: 'Système de paie en fin de vie nécessitant un remplacement sous 18 mois (~0,4 m€).' },
  { title: 'Dépendance à un prestataire unique', severity: 'Moyen', detail: 'La TMA de l\'ERP repose sur un seul prestataire sans clause de réversibilité.' },
  { title: 'Plan de reprise d\'activité', severity: 'Moyen', detail: 'PRA documenté mais non testé depuis plus de 12 mois.' },
];

export async function generate(input = {}) {
  const company = input.company || 'Société';
  const systems = input.systems?.length ? input.systems : DEFAULT_SYSTEMS;
  const findings = input.findings?.length ? input.findings : DEFAULT_FINDINGS;

  return buildReport({
    docTitle: 'Rapport de due diligence IT',
    docSubtitle: company,
    vertical: 'audit',
    docId: 'it-dd-report',
    confidential: true,
    sections: [
      {
        heading: 'Synthèse',
        blocks: [
          { type: 'p', text: `Notre revue du système d'information de ${company} couvre l'architecture applicative, l'infrastructure, la cybersécurité, l'organisation IT et les besoins d'investissement.` },
          { type: 'bullets', items: [
            `${systems.filter((s) => s.status === 'Moderne').length} systèmes cœur jugés modernes sur ${systems.length} analysés`,
            `${findings.filter((f) => f.severity === 'Élevé').length} point(s) d'attention de criticité élevée`,
            'Besoins de capex IT à intégrer au plan d\'affaires',
          ] },
        ],
      },
      {
        heading: 'Cartographie applicative',
        blocks: [
          {
            type: 'table',
            headers: ['Domaine', 'Solution', 'État', 'Commentaire'],
            rows: systems.map((s) => [s.name, s.product || '—', s.status || '—', s.comment || '—']),
          },
        ],
      },
      {
        heading: 'Cybersécurité',
        blocks: [
          { type: 'p', text: input.cyberPosture || 'Le dispositif de cybersécurité repose sur un EDR déployé sur le parc, une authentification multi-facteurs généralisée et des sauvegardes externalisées. Des tests d\'intrusion sont réalisés annuellement ; le dernier rapport ne fait pas état de vulnérabilité critique non corrigée.' },
        ],
      },
      {
        heading: 'Constats et risques',
        blocks: findings.map((f) => ({
          type: 'p',
          text: `[${f.severity}] ${f.title} — ${f.detail}`,
        })),
      },
      {
        heading: 'Besoins d\'investissement',
        blocks: [
          { type: 'bullets', items: input.capexNeeds || [
            'Remplacement du SI paie/RH : ~0,4 m€ sous 18 mois',
            'Finalisation de la migration cloud : ~0,3 m€ sur 24 mois',
            'Renforcement du PRA et tests réguliers : ~0,1 m€/an',
          ] },
        ],
      },
      {
        heading: 'Limites',
        blocks: [
          { type: 'p', text: 'Cette revue repose sur la documentation mise à disposition et les entretiens menés avec la DSI. Aucun test technique intrusif n\'a été réalisé dans le cadre de cette mission.' },
        ],
      },
    ],
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
