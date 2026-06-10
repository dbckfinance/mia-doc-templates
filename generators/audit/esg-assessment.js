// generators/audit/esg-assessment.js
// Évaluation ESG / durabilité (pptx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildDeck } from '../../shared/pptx-deck.js';

export const metadata = {
  id: 'esg-assessment',
  name: 'Évaluation ESG',
  vertical: 'audit',
  outputType: 'pptx',
  estimatedPages: '6-10 slides',
  requiredInput: ['company'],
  optionalInput: ['pillars', 'overallScore', 'regulations', 'roadmap'],
};

const DEFAULT_PILLARS = [
  {
    name: 'Environnement',
    score: 62,
    strengths: ['Bilan carbone scopes 1-2 réalisé', 'Plan de réduction énergétique engagé'],
    weaknesses: ['Scope 3 non mesuré', 'Pas d\'objectifs validés SBTi'],
  },
  {
    name: 'Social',
    score: 71,
    strengths: ['Index égalité H/F supérieur à 85', 'Faible turnover'],
    weaknesses: ['Accidentologie en hausse sur un site', 'Plan de formation à structurer'],
  },
  {
    name: 'Gouvernance',
    score: 68,
    strengths: ['Conseil avec administrateurs indépendants', 'Charte éthique déployée'],
    weaknesses: ['Pas de comité RSE formalisé', 'Dispositif anticorruption à renforcer'],
  },
];

export async function generate(input = {}) {
  const company = input.company || 'Société';
  const pillars = input.pillars?.length ? input.pillars : DEFAULT_PILLARS;
  const overall = input.overallScore ?? Math.round(pillars.reduce((s, p) => s + (p.score || 0), 0) / pillars.length);

  return buildDeck({
    title: `Évaluation ESG — ${company}`,
    subtitle: 'Diagnostic Environnement, Social, Gouvernance',
    confidential: true,
    slides: [
      {
        type: 'facts',
        title: 'Synthèse',
        facts: [
          { label: 'Score global', value: `${overall} / 100` },
          ...pillars.map((p) => ({ label: p.name, value: `${p.score} / 100` })),
        ],
      },
      {
        type: 'chart',
        title: 'Scores par pilier',
        chart: {
          kind: 'bar',
          categories: pillars.map((p) => p.name),
          series: [{ name: 'Score /100', values: pillars.map((p) => p.score || 0) }],
        },
      },
      ...pillars.map((p) => ({
        type: 'content',
        title: `${p.name} — ${p.score}/100`,
        bullets: [
          ...(p.strengths || []).map((s) => `[Force] ${s}`),
          ...(p.weaknesses || []).map((w) => `[Axe d'amélioration] ${w}`),
        ],
      })),
      {
        type: 'content',
        title: 'Conformité réglementaire',
        bullets: input.regulations || [
          'CSRD : entité dans le périmètre — premier reporting à préparer',
          'Taxonomie européenne : analyse d\'éligibilité à conduire',
          'Devoir de vigilance : seuils non atteints à date',
        ],
      },
      {
        type: 'content',
        title: 'Feuille de route recommandée',
        bullets: input.roadmap || [
          'Court terme : mesurer le scope 3 et formaliser le comité RSE',
          'Moyen terme : fixer des objectifs de réduction validés SBTi',
          'Long terme : intégrer les critères ESG à la rémunération des dirigeants',
        ],
      },
    ],
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
