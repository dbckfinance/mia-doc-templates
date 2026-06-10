// generators/audit/management-letter.js
// Lettre de recommandations à la direction (docx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildReport } from '../../shared/docx-report.js';

export const metadata = {
  id: 'management-letter',
  name: 'Lettre de Recommandations (Management Letter)',
  vertical: 'audit',
  outputType: 'docx',
  estimatedPages: '4-8',
  requiredInput: ['company'],
  optionalInput: ['fiscalYear', 'findings', 'auditor'],
};

const SEVERITY_LABELS = { high: 'Élevée', medium: 'Moyenne', low: 'Faible' };

const DEFAULT_FINDINGS = [
  {
    area: 'Cycle ventes / clients',
    severity: 'high',
    observation: 'Absence de revue indépendante des avoirs émis supérieurs à 10 k€.',
    risk: 'Risque d\'avoirs injustifiés et d\'impact sur le chiffre d\'affaires.',
    recommendation: 'Mettre en place une validation à double signature pour tout avoir supérieur à 10 k€.',
    managementResponse: 'Accepté — mise en œuvre prévue au T2.',
  },
  {
    area: 'Trésorerie',
    severity: 'medium',
    observation: 'Rapprochements bancaires non formalisés sur deux mois de l\'exercice.',
    risk: 'Risque de non-détection d\'écritures erronées ou frauduleuses.',
    recommendation: 'Formaliser mensuellement les rapprochements avec revue par le responsable comptable.',
    managementResponse: 'Accepté.',
  },
];

export async function generate(input = {}) {
  const company = input.company || 'Société';
  const fy = input.fiscalYear || new Date().getFullYear() - 1;
  const findings = input.findings?.length ? input.findings : DEFAULT_FINDINGS;

  const order = { high: 0, medium: 1, low: 2 };
  const sorted = [...findings].sort((a, b) => (order[a.severity] ?? 3) - (order[b.severity] ?? 3));

  return buildReport({
    docTitle: 'Lettre de recommandations',
    docSubtitle: `${company} — Exercice ${fy}`,
    vertical: 'audit',
    docId: 'management-letter',
    confidential: true,
    sections: [
      {
        heading: 'Introduction',
        blocks: [
          { type: 'p', text: `Dans le cadre de notre audit des comptes de ${company} pour l'exercice ${fy}, nous avons relevé certains points de contrôle interne que nous souhaitons porter à votre attention. Cette lettre ne constitue pas une revue exhaustive du contrôle interne.` },
        ],
      },
      {
        heading: 'Synthèse des constats',
        blocks: [
          {
            type: 'table',
            headers: ['#', 'Domaine', 'Criticité', 'Constat'],
            rows: sorted.map((f, i) => [String(i + 1), f.area, SEVERITY_LABELS[f.severity] || f.severity, f.observation]),
          },
        ],
      },
      ...sorted.map((f, i) => ({
        heading: `Constat ${i + 1} — ${f.area} (criticité ${SEVERITY_LABELS[f.severity] || f.severity})`,
        blocks: [
          { type: 'kv', label: 'Observation', value: f.observation },
          { type: 'kv', label: 'Risque', value: f.risk || '—' },
          { type: 'kv', label: 'Recommandation', value: f.recommendation || '—' },
          { type: 'kv', label: 'Réponse de la direction', value: f.managementResponse || 'En attente' },
        ],
      })),
      {
        heading: 'Conclusion',
        blocks: [
          { type: 'p', text: 'Nous restons à votre disposition pour échanger sur ces recommandations et leurs modalités de mise en œuvre. Nous suivrons leur avancement lors de notre prochaine intervention.' },
        ],
      },
    ],
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
