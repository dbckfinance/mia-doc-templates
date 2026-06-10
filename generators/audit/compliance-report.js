// generators/audit/compliance-report.js
// Rapport de conformité réglementaire (docx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildReport } from '../../shared/docx-report.js';

export const metadata = {
  id: 'compliance-report',
  name: 'Rapport de Conformité Réglementaire',
  vertical: 'audit',
  outputType: 'docx',
  estimatedPages: '5-10',
  requiredInput: ['company'],
  optionalInput: ['regulation', 'assessments', 'period', 'overallRating', 'actionPlan'],
};

const DEFAULT_ASSESSMENTS = [
  { requirement: 'Cartographie des risques', status: 'Conforme', comment: 'Cartographie à jour, validée par le comité des risques.' },
  { requirement: 'Dispositif de formation', status: 'Partiellement conforme', comment: 'Taux de complétion de 78% — cible 95%.' },
  { requirement: 'Procédures d\'alerte interne', status: 'Conforme', comment: 'Canal d\'alerte opérationnel et testé.' },
  { requirement: 'Contrôles de niveau 2', status: 'Non conforme', comment: 'Plan de contrôle non formalisé sur le périmètre international.' },
];

export async function generate(input = {}) {
  const company = input.company || 'Société';
  const regulation = input.regulation || 'Dispositif Sapin II / AML-CFT';
  const assessments = input.assessments?.length ? input.assessments : DEFAULT_ASSESSMENTS;

  const conformCount = assessments.filter((a) => /^conforme/i.test(a.status)).length;

  return buildReport({
    docTitle: 'Rapport de conformité réglementaire',
    docSubtitle: `${company} — ${regulation}`,
    vertical: 'audit',
    docId: 'compliance-report',
    confidential: true,
    sections: [
      {
        heading: 'Synthèse',
        blocks: [
          { type: 'kv', label: 'Référentiel évalué', value: regulation },
          { type: 'kv', label: 'Période', value: input.period || 'Exercice en cours' },
          { type: 'kv', label: 'Exigences conformes', value: `${conformCount} / ${assessments.length}` },
          { type: 'kv', label: 'Appréciation globale', value: input.overallRating || (conformCount === assessments.length ? 'Conforme' : 'Conformité partielle — plan d\'action requis') },
        ],
      },
      {
        heading: 'Évaluation détaillée',
        blocks: [
          {
            type: 'table',
            headers: ['Exigence', 'Statut', 'Commentaire'],
            rows: assessments.map((a) => [a.requirement, a.status, a.comment || '—']),
          },
        ],
      },
      {
        heading: 'Plan d\'action',
        blocks: [
          { type: 'bullets', items: input.actionPlan || assessments
            .filter((a) => !/^conforme$/i.test(a.status))
            .map((a) => `${a.requirement} : remédiation à engager — ${a.comment || 'définir le plan avec le responsable conformité'}`) },
        ],
      },
      {
        heading: 'Limites de la revue',
        blocks: [
          { type: 'p', text: 'Cette revue de conformité repose sur les informations et documents communiqués à la date du rapport. Elle ne constitue pas une garantie d\'absence de manquement et ne se substitue pas à un avis juridique.' },
        ],
      },
    ],
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
