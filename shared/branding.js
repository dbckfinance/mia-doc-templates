// shared/branding.js
// M&IA brand constants shared by every generator.
// Colors are hex WITHOUT the leading '#' (docx/exceljs convention);
// pptx helpers add the '#' where needed.

export const BRAND = {
  // Core palette
  navy: '1B2A4A',
  darkBlue: '2C3E6B',
  accent: '4472C4',
  lightBlue: 'D6E4F0',
  lightGray: 'F5F5F5',
  medGray: '808080',
  text: '333333',
  white: 'FFFFFF',
  positive: '548235',
  negative: 'C00000',
  warning: 'BF8F00',

  // Typography
  font: 'Calibri',
  fontSerif: 'Garamond',

  // Marks
  confidentiality: 'Strictly Confidential',
  confidentialityFr: 'Strictement Confidentiel',
  brandName: 'M&IA',
  disclaimerShort:
    'This document has been prepared by M&IA for discussion purposes only and does not constitute an offer or solicitation.',
};

/** Convention per vertical for exhibit labelling. */
export const EXHIBIT_LABELS = {
  ma: 'Exhibit',
  ib: 'Exhibit',
  audit: 'Appendix',
  'hedge-fund': 'Exhibit',
  'equity-research': 'Figure',
  common: 'Annex',
};

/** '#'-prefixed color for pptxgenjs. */
export function pptxColor(hex) {
  return hex.startsWith('#') ? hex : hex;
}

/** ARGB color for exceljs fills. */
export function argb(hex) {
  return `FF${hex.replace('#', '')}`;
}
