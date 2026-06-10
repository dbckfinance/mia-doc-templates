// shared/disclaimer-library.js
// Legal disclaimers per document type / vertical. French by default,
// English variants where market convention requires.

export const DISCLAIMERS = {
  generic:
    "Ce document a été préparé par M&IA à des fins de discussion uniquement. Il ne constitue ni une offre, ni une sollicitation d'achat ou de vente de titres. Les informations contenues dans ce document proviennent de sources jugées fiables mais ne font l'objet d'aucune garantie quant à leur exactitude ou exhaustivité.",

  ma:
    "Ce document est strictement confidentiel et destiné exclusivement à son destinataire. Il ne peut être reproduit ou distribué sans l'accord écrit préalable de M&IA. Les projections financières présentées reposent sur des hypothèses qui peuvent ne pas se réaliser. M&IA n'assume aucune responsabilité quant aux décisions prises sur la base de ce document.",

  ib:
    'This material is for institutional clients only and is not intended for retail distribution. It does not constitute investment advice, an offer, or a solicitation. Past performance is not indicative of future results. Securities transactions involve risk, including possible loss of principal.',

  audit:
    "Ce rapport a été établi dans le cadre de la mission décrite dans la lettre de mission et conformément aux normes d'exercice professionnel applicables. Il est destiné exclusivement à l'usage de la direction et des organes de gouvernance de l'entité, et ne peut être communiqué à des tiers sans accord préalable.",

  'hedge-fund':
    'This document is provided to qualified investors on a confidential basis. It does not constitute an offer to sell or a solicitation to buy interests in any fund. An investment in the fund involves substantial risk, including loss of capital. Performance figures are unaudited and subject to revision.',

  'equity-research':
    "Cette note de recherche est produite à titre d'information uniquement et ne constitue pas un conseil en investissement personnalisé. Les opinions exprimées reflètent le jugement de l'analyste à la date de publication et sont susceptibles d'évoluer sans préavis. L'investisseur est seul responsable de ses décisions d'investissement.",

  legal:
    "Ce projet de document est fourni à titre indicatif et ne constitue pas un conseil juridique. Il doit impérativement être revu et adapté par un conseil juridique qualifié avant toute signature ou utilisation.",
};

export function disclaimerFor(vertical, docType) {
  if (['nda', 'loi', 'process-letter', 'engagement-letter', 'capital-call-notice'].includes(docType)) {
    return DISCLAIMERS.legal;
  }
  return DISCLAIMERS[vertical] || DISCLAIMERS.generic;
}
