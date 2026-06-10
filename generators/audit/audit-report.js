// generators/audit/audit-report.js
// Rapport d'audit légal (docx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildReport } from '../../shared/docx-report.js';

export const metadata = {
  id: 'audit-report',
  name: "Rapport d'Audit Légal",
  vertical: 'audit',
  outputType: 'docx',
  estimatedPages: '5-10',
  requiredInput: ['company'],
  optionalInput: ['fiscalYear', 'opinion', 'keyAuditMatters', 'auditor', 'emphasisOfMatter'],
};

const OPINIONS = {
  unqualified: 'certification sans réserve',
  qualified: 'certification avec réserve(s)',
  adverse: 'refus de certifier',
  disclaimer: 'impossibilité de certifier',
};

export async function generate(input = {}) {
  const company = input.company || 'Société';
  const fy = input.fiscalYear || new Date().getFullYear() - 1;
  const opinionKey = input.opinion || 'unqualified';
  const auditor = input.auditor || 'M&IA Audit';
  const kams = input.keyAuditMatters || [
    { title: 'Reconnaissance du revenu', description: 'Risque lié au cut-off et aux contrats à éléments multiples.', response: 'Tests de détail sur les contrats significatifs, revue du cut-off de clôture.' },
    { title: 'Évaluation des écarts d\'acquisition', description: 'Sensibilité des tests de dépréciation aux hypothèses de taux et de flux.', response: 'Revue des modèles de dépréciation, contre-expertise des hypothèses clés.' },
  ];

  return buildReport({
    docTitle: `Rapport du commissaire aux comptes — Exercice ${fy}`,
    docSubtitle: company,
    vertical: 'audit',
    docId: 'audit-report',
    confidential: false,
    sections: [
      {
        heading: 'Opinion',
        blocks: [
          { type: 'p', text: `En exécution de la mission qui nous a été confiée, nous avons effectué l'audit des comptes annuels de la société ${company} relatifs à l'exercice clos le 31 décembre ${fy}.` },
          { type: 'p', text: `Nous certifions que les comptes annuels sont, au regard des règles et principes comptables français, réguliers et sincères et donnent une image fidèle du résultat des opérations de l'exercice écoulé ainsi que de la situation financière et du patrimoine de la société à la fin de cet exercice. Conclusion : ${OPINIONS[opinionKey] || OPINIONS.unqualified}.` },
        ],
      },
      {
        heading: 'Fondement de l\'opinion',
        blocks: [
          { type: 'p', text: 'Nous avons effectué notre audit selon les normes d\'exercice professionnel applicables en France. Nous estimons que les éléments que nous avons collectés sont suffisants et appropriés pour fonder notre opinion.' },
          { type: 'p', text: 'Les responsabilités qui nous incombent en vertu de ces normes sont indiquées dans la partie « Responsabilités du commissaire aux comptes relatives à l\'audit des comptes annuels » du présent rapport.' },
        ],
      },
      ...(input.emphasisOfMatter ? [{
        heading: 'Observation',
        blocks: [{ type: 'p', text: input.emphasisOfMatter }],
      }] : []),
      {
        heading: 'Justification des appréciations — Points clés de l\'audit',
        blocks: kams.flatMap((k) => [
          { type: 'h2', text: k.title },
          { type: 'p', text: `Risque identifié : ${k.description}` },
          { type: 'p', text: `Réponse d'audit : ${k.response}` },
        ]),
      },
      {
        heading: 'Responsabilités de la direction',
        blocks: [
          { type: 'p', text: 'Il appartient à la direction d\'établir des comptes annuels présentant une image fidèle conformément aux règles et principes comptables français ainsi que de mettre en place le contrôle interne qu\'elle estime nécessaire à l\'établissement de comptes annuels ne comportant pas d\'anomalies significatives, que celles-ci proviennent de fraudes ou résultent d\'erreurs.' },
        ],
      },
      {
        heading: 'Responsabilités du commissaire aux comptes',
        blocks: [
          { type: 'p', text: 'Il nous appartient d\'établir un rapport sur les comptes annuels. Notre objectif est d\'obtenir l\'assurance raisonnable que les comptes annuels pris dans leur ensemble ne comportent pas d\'anomalies significatives.' },
          { type: 'bullets', items: [
            'Identification et évaluation des risques d\'anomalies significatives',
            'Prise de connaissance du contrôle interne pertinent pour l\'audit',
            'Appréciation du caractère approprié des méthodes comptables retenues',
            'Appréciation de la présentation d\'ensemble des comptes annuels',
          ] },
        ],
      },
      {
        heading: 'Signature',
        blocks: [
          { type: 'kv', label: 'Cabinet', value: auditor },
          { type: 'kv', label: 'Date', value: input.signatureDate || new Date().toLocaleDateString('fr-FR') },
          { type: 'kv', label: 'Lieu', value: input.signatureCity || 'Paris' },
        ],
      },
    ],
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
