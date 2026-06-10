// generators/equity-research/initiation-coverage.js
// Initiation de couverture (docx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildReport } from '../../shared/docx-report.js';

export const metadata = {
  id: 'initiation-coverage',
  name: 'Initiation de Couverture',
  vertical: 'equity-research',
  outputType: 'docx',
  estimatedPages: '20-40',
  requiredInput: ['ticker'],
  optionalInput: ['company', 'rating', 'targetPrice', 'currentPrice', 'sector', 'thesis', 'financials', 'valuation', 'risks', 'analyst'],
};

export async function generate(input = {}) {
  const ticker = input.ticker || 'TICKER';
  const company = input.company || ticker;
  const rating = input.rating || 'Achat';
  const target = input.targetPrice;
  const current = input.currentPrice;
  const upside = current && target ? `${((target / current - 1) * 100).toFixed(1)}%` : 'n.d.';

  return buildReport({
    docTitle: `Initiation de couverture — ${company}`,
    docSubtitle: `${ticker} | ${rating} | Objectif : ${target ?? 'n.d.'} (${upside})`,
    vertical: 'equity-research',
    docId: 'initiation-coverage',
    confidential: false,
    toc: true,
    sections: [
      {
        heading: 'Synthèse de la recommandation',
        blocks: [
          { type: 'facts', facts: [
            { label: 'Recommandation', value: rating },
            { label: 'Objectif de cours', value: target != null ? String(target) : 'n.d.' },
            { label: 'Cours actuel', value: current != null ? String(current) : 'n.d.' },
            { label: 'Potentiel', value: upside },
            { label: 'Secteur', value: input.sector || '—' },
            { label: 'Analyste', value: input.analyst || 'M&IA Research' },
          ] },
          { type: 'bullets', items: input.thesis || [
            'Positionnement de leader sur un marché en croissance structurelle',
            'Levier opérationnel sous-estimé par le consensus',
            'Valorisation attractive au regard du profil croissance/rentabilité',
          ] },
        ],
      },
      {
        heading: 'Présentation de la société',
        blocks: [
          { type: 'p', text: input.companyOverview || `${company} est un acteur de référence de son marché. Cette section présente le modèle économique, les segments d'activité, la répartition géographique du chiffre d'affaires et l'historique du groupe.` },
        ],
      },
      {
        heading: 'Analyse du marché et positionnement concurrentiel',
        blocks: [
          { type: 'p', text: input.marketAnalysis || 'Analyse de la taille de marché, des moteurs de croissance, de l\'intensité concurrentielle et des barrières à l\'entrée. Le positionnement relatif de la société y est évalué (parts de marché, avantages compétitifs, pricing power).' },
        ],
      },
      {
        heading: 'Prévisions financières',
        blocks: input.financials?.rows?.length
          ? [{ type: 'table', headers: input.financials.headers || ['Indicateur', 'N-1', 'N', 'N+1e', 'N+2e'], rows: input.financials.rows }]
          : [{ type: 'p', text: 'Nos estimations détaillées (compte de résultat, flux de trésorerie, bilan) figurent dans le modèle joint. Les hypothèses clés portent sur la croissance organique, l\'évolution des marges et la conversion en cash.' }],
      },
      {
        heading: 'Valorisation',
        blocks: input.valuation?.rows?.length
          ? [{ type: 'table', headers: input.valuation.headers || ['Méthode', 'Poids', 'Valeur par action'], rows: input.valuation.rows }]
          : [{ type: 'p', text: 'Notre objectif de cours résulte d\'une approche multicritère combinant DCF, comparables boursiers et transactions de référence, pondérés selon leur pertinence pour le profil de la société.' }],
      },
      {
        heading: 'Risques sur la recommandation',
        blocks: [
          { type: 'bullets', items: input.risks || [
            'Sensibilité au cycle macroéconomique',
            'Risque d\'exécution sur les initiatives stratégiques',
            'Pression concurrentielle accrue sur les prix',
            'Risques réglementaires propres au secteur',
          ] },
        ],
      },
    ],
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
