// generators/audit/vendor-dd-report.js
// Rapport de Vendor Due Diligence (docx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildReport } from '../../shared/docx-report.js';

export const metadata = {
  id: 'vendor-dd-report',
  name: 'Rapport Vendor Due Diligence',
  vertical: 'audit',
  outputType: 'docx',
  estimatedPages: '15-30',
  requiredInput: ['company'],
  optionalInput: ['sector', 'period', 'keyFindings', 'financialHighlights', 'risks'],
};

export async function generate(input = {}) {
  const company = input.company || 'Société';
  const sector = input.sector || 'son secteur';

  return buildReport({
    docTitle: 'Vendor Due Diligence — Rapport financier',
    docSubtitle: company,
    vertical: 'audit',
    docId: 'vendor-dd-report',
    confidential: true,
    toc: true,
    sections: [
      {
        heading: 'Résumé exécutif',
        blocks: [
          { type: 'p', text: input.executiveSummary || `Ce rapport de vendor due diligence présente une analyse indépendante de la performance financière historique de ${company}, acteur de ${sector}, en vue de sa cession. Il couvre la qualité des résultats, le BFR, la dette nette et les principaux risques financiers.` },
          { type: 'bullets', items: input.keyFindings || [
            'Croissance du chiffre d\'affaires régulière sur la période analysée',
            'EBITDA ajusté retraité des éléments non récurrents identifiés',
            'BFR saisonnier nécessitant un mécanisme d\'ajustement de prix adapté',
            'Aucun passif hors bilan significatif identifié',
          ] },
        ],
      },
      {
        heading: 'Périmètre et démarche',
        blocks: [
          { type: 'kv', label: 'Société analysée', value: company },
          { type: 'kv', label: 'Période couverte', value: input.period || '3 derniers exercices + situation intermédiaire' },
          { type: 'bullets', items: [
            'Analyse des comptes annuels et des reportings de gestion',
            'Entretiens avec la direction financière',
            'Revue de la qualité des résultats (QoE) et du BFR normatif',
            'Analyse de la dette nette et des engagements hors bilan',
          ] },
        ],
      },
      {
        heading: 'Performance financière historique',
        blocks: input.financialHighlights?.length
          ? [{ type: 'table', headers: ['Indicateur', ...(input.financialHighlights[0]?.values?.map((_, i) => `A-${input.financialHighlights[0].values.length - 1 - i}`) || [])], rows: input.financialHighlights.map((h) => [h.label, ...(h.values || [])]) }]
          : [{ type: 'p', text: 'Les agrégats financiers détaillés (chiffre d\'affaires, marges, EBITDA, conversion en cash) sont présentés dans les annexes chiffrées et le databook joint.' }],
      },
      {
        heading: 'Qualité des résultats',
        blocks: [
          { type: 'p', text: 'Nos travaux de qualité des résultats ont consisté à identifier les éléments non récurrents, les normalisations de charges et produits, et les éléments pro forma. Le détail figure dans le pont EBITDA du databook.' },
        ],
      },
      {
        heading: 'BFR et dette nette',
        blocks: [
          { type: 'p', text: 'Le BFR normatif a été déterminé sur la base des moyennes mensuelles, en tenant compte de la saisonnalité. Les éléments assimilables à de la dette ont été identifiés et chiffrés.' },
        ],
      },
      {
        heading: 'Points d\'attention',
        blocks: [
          { type: 'bullets', items: input.risks || [
            'Concentration client à surveiller dans la durée',
            'Dépendance à certains fournisseurs stratégiques',
            'Investissements de maintenance à anticiper',
          ] },
        ],
      },
      {
        heading: 'Limites et diffusion',
        blocks: [
          { type: 'p', text: 'Ce rapport est destiné aux acquéreurs potentiels dans le cadre du processus de cession, sous réserve de la signature d\'une lettre d\'accès (release letter). Il ne constitue ni un audit ni une garantie.' },
        ],
      },
    ],
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
