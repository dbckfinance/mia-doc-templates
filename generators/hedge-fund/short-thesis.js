// generators/hedge-fund/short-thesis.js
// Thèse short (docx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildReport } from '../../shared/docx-report.js';

export const metadata = {
  id: 'short-thesis',
  name: 'Thèse Short',
  vertical: 'hedge-fund',
  outputType: 'docx',
  estimatedPages: '6-12',
  requiredInput: ['ticker'],
  optionalInput: ['company', 'currentPrice', 'targetPrice', 'redFlags', 'catalysts', 'risks', 'borrowCost', 'shortInterest'],
};

export async function generate(input = {}) {
  const ticker = input.ticker || 'TICKER';
  const company = input.company || ticker;
  const current = input.currentPrice;
  const target = input.targetPrice;
  const downside = current && target ? `${((target / current - 1) * 100).toFixed(1)}%` : 'n.d.';

  return buildReport({
    docTitle: `Thèse short — ${company} (${ticker})`,
    docSubtitle: `Objectif ${target ?? 'n.d.'} (${downside})`,
    vertical: 'hedge-fund',
    docId: 'short-thesis',
    confidential: true,
    sections: [
      {
        heading: 'Résumé',
        blocks: [
          { type: 'facts', facts: [
            { label: 'Cours actuel', value: current != null ? String(current) : 'n.d.' },
            { label: 'Objectif de cours', value: target != null ? String(target) : 'n.d.' },
            { label: 'Potentiel de baisse', value: downside },
            { label: 'Coût d\'emprunt', value: input.borrowCost || 'n.d.' },
            { label: 'Short interest', value: input.shortInterest || 'n.d.' },
          ] },
        ],
      },
      {
        heading: 'Signaux d\'alerte (red flags)',
        blocks: [
          { type: 'bullets', items: input.redFlags || [
            'Divergence croissante entre résultat comptable et génération de cash',
            'Rotation élevée du management financier et des auditeurs',
            'Hausse anormale des créances clients vs chiffre d\'affaires',
            'Acquisitions séquentielles masquant la décroissance organique',
            'Ventes d\'initiés significatives',
          ] },
        ],
      },
      {
        heading: 'Analyse fondamentale',
        blocks: [
          { type: 'p', text: input.fundamentalAnalysis || 'La dégradation des fondamentaux (pression sur les marges, perte de parts de marché, levier croissant) n\'est pas reflétée dans la valorisation actuelle, qui intègre un scénario de croissance jugé irréaliste.' },
        ],
      },
      {
        heading: 'Catalyseurs',
        blocks: [
          { type: 'bullets', items: input.catalysts || [
            'Avertissement sur résultats attendu sous 2 trimestres',
            'Refinancement de la dette dans des conditions dégradées',
            'Expiration de lock-up / cession d\'un actionnaire de référence',
          ] },
        ],
      },
      {
        heading: 'Risques de la position short',
        blocks: [
          { type: 'bullets', items: input.risks || [
            'Risque de squeeze (short interest élevé, flottant réduit)',
            'OPA / soutien d\'un actionnaire stratégique',
            'Rachat d\'actions massif soutenant le cours',
            'Coût de portage si l\'horizon s\'allonge',
          ] },
        ],
      },
      {
        heading: 'Gestion de la position',
        blocks: [
          { type: 'kv', label: 'Taille initiale', value: input.positionSize || '1,5% NAV' },
          { type: 'kv', label: 'Stop de réévaluation', value: input.stopLoss || '+20% vs prix d\'entrée' },
          { type: 'kv', label: 'Couverture', value: input.hedge || 'Calls de couverture hors de la monnaie si short interest > 15%' },
        ],
      },
    ],
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
