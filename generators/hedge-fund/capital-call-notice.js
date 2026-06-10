// generators/hedge-fund/capital-call-notice.js
// Avis d'appel de fonds / distribution (docx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildLetter } from '../../shared/docx-report.js';

export const metadata = {
  id: 'capital-call-notice',
  name: "Avis d'Appel de Fonds / Distribution",
  vertical: 'hedge-fund',
  outputType: 'docx',
  estimatedPages: '1-2',
  requiredInput: ['fund', 'investor'],
  optionalInput: ['noticeType', 'amount', 'currency', 'dueDate', 'commitment', 'calledToDate', 'purpose', 'bankDetails'],
};

export async function generate(input = {}) {
  const fund = input.fund || 'Fonds';
  const investor = input.investor || 'Investisseur';
  const type = (input.noticeType || 'call').toLowerCase();
  const isCall = type !== 'distribution';
  const amount = input.amount ?? 500000;
  const currency = input.currency || 'EUR';
  const fmtAmount = `${amount.toLocaleString('fr-FR')} ${currency}`;

  const paragraphs = isCall
    ? [
        `Conformément aux dispositions du règlement du fonds ${fund}, nous procédons par la présente à un appel de fonds.`,
        `Montant appelé : ${fmtAmount}, à régler au plus tard le ${input.dueDate || 'dans 10 jours ouvrés'}.`,
        ...(input.commitment != null ? [`Votre engagement total s'élève à ${Number(input.commitment).toLocaleString('fr-FR')} ${currency}, dont ${input.calledToDate != null ? Number(input.calledToDate).toLocaleString('fr-FR') : '—'} ${currency} appelés à ce jour (incluant le présent appel).`] : []),
        `Objet de l'appel : ${input.purpose || 'financement des investissements en cours et des frais du fonds'}.`,
        `Le règlement doit être effectué par virement sur le compte suivant : ${input.bankDetails || '[coordonnées bancaires communiquées séparément par sécurité]'}.`,
        'En cas de retard de paiement, les dispositions du règlement relatives aux investisseurs défaillants trouveront à s\'appliquer.',
      ]
    : [
        `Nous avons le plaisir de vous informer qu'une distribution du fonds ${fund} sera effectuée à votre profit.`,
        `Montant distribué : ${fmtAmount}, avec date de valeur le ${input.dueDate || 'prochain jour ouvré'}.`,
        `Origine de la distribution : ${input.purpose || 'produit de cession d\'investissements du portefeuille'}.`,
        'Le versement sera effectué sur le compte bancaire enregistré dans nos livres. Nous vous invitons à nous signaler tout changement de coordonnées bancaires.',
        'Un état récapitulatif de votre situation (capital appelé, distribué, NAV résiduelle) est joint au présent avis.',
      ];

  return buildLetter({
    docTitle: isCall ? `Avis d'appel de fonds — ${fund}` : `Avis de distribution — ${fund}`,
    docSubtitle: `À l'attention de ${investor}`,
    vertical: 'hedge-fund',
    docId: 'capital-call-notice',
    confidential: true,
    recipient: investor,
    paragraphs,
    signature: {
      name: input.signatory || 'La société de gestion',
      title: fund,
    },
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
