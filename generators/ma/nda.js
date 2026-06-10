// generators/ma/nda.js
// Accord de confidentialité / NDA (docx) — draft à faire valider par un juriste

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildReport } from '../../shared/docx-report.js';
import { disclaimerFor } from '../../shared/disclaimer-library.js';

export const metadata = {
  id: 'nda',
  name: 'Accord de Confidentialité (NDA)',
  vertical: 'ma',
  outputType: 'docx',
  estimatedPages: '4-8',
  requiredInput: ['disclosingParty', 'receivingParty'],
  optionalInput: ['project', 'duration', 'jurisdiction', 'purpose', 'extraClauses'],
};

export async function generate(input = {}) {
  const disclosing = input.disclosingParty || '[Partie Divulgatrice]';
  const receiving = input.receivingParty || '[Partie Réceptrice]';
  const duration = input.duration || '24 mois';
  const jurisdiction = input.jurisdiction || 'tribunaux de Paris, droit français';
  const purpose = input.purpose || `l'étude d'une éventuelle opération de rapprochement (le « Projet »)`;

  const sections = [
    {
      heading: 'Parties',
      paragraphs: [
        `Le présent accord de confidentialité (l'« Accord ») est conclu entre ${disclosing} (la « Partie Divulgatrice ») et ${receiving} (la « Partie Réceptrice »), ci-après ensemble les « Parties ».`,
      ],
    },
    {
      heading: 'Article 1 — Objet',
      paragraphs: [
        `Dans le cadre de ${purpose}, la Partie Divulgatrice est susceptible de communiquer à la Partie Réceptrice des informations confidentielles. Le présent Accord définit les conditions dans lesquelles ces informations sont communiquées et protégées.`,
      ],
    },
    {
      heading: 'Article 2 — Informations Confidentielles',
      paragraphs: [
        `Sont considérées comme confidentielles toutes informations, de quelque nature qu'elles soient (commerciale, financière, technique, juridique, stratégique), communiquées par la Partie Divulgatrice, par écrit ou oralement, ainsi que l'existence même du Projet et des discussions entre les Parties.`,
      ],
      bullets: [
        'Ne sont pas confidentielles les informations déjà publiques sans violation du présent Accord',
        'Ne sont pas confidentielles les informations déjà connues de la Partie Réceptrice avant communication',
        'Ne sont pas confidentielles les informations développées de manière indépendante',
        "Ne sont pas confidentielles les informations reçues d'un tiers non tenu à une obligation de confidentialité",
      ],
    },
    {
      heading: 'Article 3 — Engagements de la Partie Réceptrice',
      bullets: [
        'Préserver la stricte confidentialité des Informations Confidentielles',
        "N'utiliser les Informations Confidentielles qu'aux seules fins de l'évaluation du Projet",
        "Ne divulguer les Informations Confidentielles qu'aux collaborateurs et conseils ayant besoin d'en connaître, eux-mêmes tenus à des obligations équivalentes",
        'Ne pas contacter les clients, fournisseurs, ou salariés de la Partie Divulgatrice sans accord préalable écrit',
        'Restituer ou détruire les Informations Confidentielles à première demande',
      ],
    },
    {
      heading: 'Article 4 — Non-sollicitation',
      paragraphs: [
        `Pendant la durée du présent Accord et pour une période de douze (12) mois suivant son expiration, la Partie Réceptrice s'interdit de solliciter ou d'embaucher, directement ou indirectement, tout salarié clé de la Partie Divulgatrice rencontré dans le cadre du Projet.`,
      ],
    },
    {
      heading: 'Article 5 — Durée',
      paragraphs: [
        `Le présent Accord entre en vigueur à la date de sa signature et demeure applicable pendant une durée de ${duration}. Les obligations de confidentialité survivent à son expiration pour les informations constituant un secret des affaires.`,
      ],
    },
    {
      heading: 'Article 6 — Absence d\'engagement',
      paragraphs: [
        `Le présent Accord ne constitue en aucun cas un engagement de conclure une quelconque opération. Chaque Partie demeure libre de mettre fin aux discussions à tout moment, sans indemnité.`,
      ],
    },
    {
      heading: 'Article 7 — Droit applicable et juridiction',
      paragraphs: [
        `Le présent Accord est soumis au droit applicable et aux juridictions suivantes : ${jurisdiction}.`,
      ],
    },
  ];

  for (const clause of input.extraClauses || []) {
    sections.push({ heading: clause.heading, paragraphs: clause.paragraphs || [], bullets: clause.bullets || [] });
  }

  sections.push({
    heading: 'Signatures',
    paragraphs: [
      `Fait en deux exemplaires originaux, le ${input.date || '____________'}.`,
      '',
      `Pour ${disclosing} :  _______________________`,
      '',
      `Pour ${receiving} :  _______________________`,
    ],
  });

  return buildReport({
    docTitle: 'Accord de Confidentialité',
    docSubtitle: input.project || 'Projet confidentiel',
    date: input.date,
    lang: 'fr',
    includeToc: false,
    sections,
    disclaimer: disclaimerFor('ma', metadata.id),
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
