// generators/audit/internal-audit-memo.js
// Mémo d'audit interne (docx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildReport } from '../../shared/docx-report.js';

export const metadata = {
  id: 'internal-audit-memo',
  name: "Mémo d'Audit Interne",
  vertical: 'audit',
  outputType: 'docx',
  estimatedPages: '3-6',
  requiredInput: ['subject'],
  optionalInput: ['company', 'scope', 'methodology', 'findings', 'conclusion', 'auditor', 'period'],
};

export async function generate(input = {}) {
  const subject = input.subject || 'Processus audité';
  const company = input.company || 'Société';
  const findings = input.findings || [
    { title: 'Séparation des tâches', detail: 'Le même collaborateur saisit et valide les paiements fournisseurs.', rating: 'Majeur' },
    { title: 'Documentation des procédures', detail: 'Procédures non mises à jour depuis plus de 24 mois.', rating: 'Modéré' },
  ];

  return buildReport({
    docTitle: `Mémo d'audit interne — ${subject}`,
    docSubtitle: company,
    vertical: 'audit',
    docId: 'internal-audit-memo',
    confidential: true,
    sections: [
      {
        heading: 'Contexte et objectifs',
        blocks: [
          { type: 'kv', label: 'Entité', value: company },
          { type: 'kv', label: 'Processus audité', value: subject },
          { type: 'kv', label: 'Période couverte', value: input.period || 'Exercice en cours' },
          { type: 'kv', label: 'Auditeur', value: input.auditor || 'Audit interne M&IA' },
          { type: 'p', text: input.scope || `La mission a porté sur l'évaluation de l'efficacité du dispositif de contrôle interne relatif au processus « ${subject} ».` },
        ],
      },
      {
        heading: 'Méthodologie',
        blocks: [
          { type: 'bullets', items: input.methodology || [
            'Entretiens avec les opérationnels et responsables du processus',
            'Revue documentaire (procédures, organigrammes, délégations)',
            'Tests de cheminement (walkthrough) et tests de détail',
            'Analyse de données sur la période auditée',
          ] },
        ],
      },
      {
        heading: 'Constats',
        blocks: [
          {
            type: 'table',
            headers: ['#', 'Constat', 'Cotation'],
            rows: findings.map((f, i) => [String(i + 1), f.title, f.rating || '—']),
          },
          ...findings.map((f, i) => ({ type: 'p', text: `${i + 1}. ${f.title} : ${f.detail}` })),
        ],
      },
      {
        heading: 'Conclusion et plan d\'action',
        blocks: [
          { type: 'p', text: input.conclusion || 'Le dispositif de contrôle interne du processus audité est globalement perfectible. Les constats ci-dessus feront l\'objet d\'un plan d\'action suivi par l\'audit interne avec des échéances trimestrielles.' },
        ],
      },
    ],
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
