/**
 * Génère une présentation IB-style : multiples SaaS (M&A récents).
 * Sortie : artifacts/Presentation_Deals_SaaS.pptx
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import PptxGenJS from 'pptxgenjs';
import { BRAND } from '../../shared/branding.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(__dirname, '../..');
const artifactsDir = path.join(repoRoot, 'artifacts');
const outFile = path.join(artifactsDir, 'Presentation_Deals_SaaS.pptx');

fs.mkdirSync(artifactsDir, { recursive: true });

const PRIMARY = BRAND.primary.replace('#', '');
const ACCENT = BRAND.accent.replace('#', '');
const TEXT = '333333';
const MUTED = '64748B';
const HEADER_BG = PRIMARY;
const HEADER_TX = 'FFFFFF';
const ROW_ALT = 'F8FAFC';

const deals = [
  {
    cible: 'Perfect Corp',
    acquereur: 'Cyberlink Corp',
    geo: 'Taïwan',
    ev: 198.6,
    ebitda: 0.67,
    evEbitda: 297.8,
    evRev: 2.9,
  },
  {
    cible: 'Alia Software Inc',
    acquereur: 'dotDigital Group plc',
    geo: 'États-Unis',
    ev: 60.0,
    ebitda: 1.0,
    evEbitda: 60.0,
    evRev: 15.0,
  },
  {
    cible: 'Infomedia Ltd',
    acquereur: 'TPG Capital LP',
    geo: 'Australie',
    ev: 382.8,
    ebitda: 19.0,
    evEbitda: 20.2,
    evRev: 4.1,
  },
];

const medianEvEbitda = 60.0;
const medianEvRev = 4.1;

const fmtM = (n) =>
  n.toLocaleString('fr-FR', { minimumFractionDigits: 1, maximumFractionDigits: 1 });

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'M&IA Document Templates';
pptx.title = 'Analyse des Multiples — SaaS';
pptx.subject = 'Multiples de valorisation';

function addDividerLine(slide, y = 1.05) {
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.5,
    y,
    w: 12.33,
    h: 0.02,
    fill: { color: ACCENT },
    line: { color: ACCENT, width: 0 },
  });
}

// --- Slide 1 : Titre ---
const s1 = pptx.addSlide();
s1.background = { color: PRIMARY };
s1.addText('Analyse des Multiples de Valorisation - Secteur SaaS', {
  x: 0.6,
  y: 2.35,
  w: 12.2,
  h: 1.35,
  fontSize: 32,
  bold: true,
  color: 'FFFFFF',
  align: 'center',
  fontFace: 'Arial',
});
s1.addText('Transactions récentes — Comparaison EV / EBITDA / Revenus', {
  x: 0.6,
  y: 3.85,
  w: 12.2,
  h: 0.55,
  fontSize: 14,
  color: 'CBD5E1',
  align: 'center',
  fontFace: 'Arial',
});
s1.addText(`${BRAND.confidentiality} · ${new Date().toLocaleDateString('fr-FR', { year: 'numeric', month: 'long', day: 'numeric' })}`, {
  x: 0.6,
  y: 5.85,
  w: 12.2,
  h: 0.45,
  fontSize: 11,
  color: '94A3B8',
  align: 'center',
  fontFace: 'Arial',
});

// --- Slide 2 : Executive Summary ---
const s2 = pptx.addSlide();
s2.addText('Executive Summary', {
  x: 0.5,
  y: 0.35,
  w: 12,
  h: 0.65,
  fontSize: 26,
  bold: true,
  color: PRIMARY,
  fontFace: 'Arial',
});
addDividerLine(s2, 1.0);
s2.addText(
  [
    { text: 'Tendances clés\n', options: { breakLine: true, bold: true, fontSize: 14, color: PRIMARY } },
    {
      text:
        '• Échantillon de trois opérations SaaS / marketing tech récentes avec des profils géographiques et d’acquéreurs hétérogènes (Taïwan, États-Unis, Australie).\n',
    },
    {
      text:
        '• Forte dispersion des multiples EV/EBITDA (fourchette observée ~20x à ~298x), reflétant la sensibilité au niveau d’EBITDA normalisé et aux narratifs de croissance.\n',
    },
    {
      text:
        '• Les multiples EV/Revenus sont matériellement plus comparables pour certaines cibles à faible maturité EBITDA (à nuancer par la qualité du rev recurring [à compléter]).\n',
    },
    { text: '\nKPI — Médianes (échantillon n=3)\n', options: { breakLine: true, bold: true, fontSize: 14, color: PRIMARY } },
    { text: `• Médiane EV/EBITDA : ${fmtM(medianEvEbitda)}x\n` },
    { text: `• Médiane EV/Revenus : ${fmtM(medianEvRev)}x\n` },
    {
      text:
        '\nSources : communiqués de presse, documentation publique des sociétés et estimations internes [à compléter pour liens exacts].',
      options: { italic: true, fontSize: 11, color: MUTED },
    },
  ],
  {
    x: 0.55,
    y: 1.2,
    w: 12.2,
    h: 5.5,
    fontSize: 13,
    color: TEXT,
    fontFace: 'Arial',
    valign: 'top',
    lineSpacingMultiple: 1.15,
  }
);

// --- Slide 3 : Tableau des transactions ---
const s3 = pptx.addSlide();
s3.addText('Vue Transactionnelle — Détail des trois deals', {
  x: 0.5,
  y: 0.35,
  w: 12,
  h: 0.65,
  fontSize: 22,
  bold: true,
  color: PRIMARY,
  fontFace: 'Arial',
});
addDividerLine(s3, 1.0);

const headerOpts = {
  bold: true,
  fontFace: 'Arial',
  fontSize: 10,
  color: HEADER_TX,
  fill: { color: HEADER_BG },
  valign: 'middle',
  align: 'left',
};

const cell = (text, alt = false, align = 'left') => ({
  text: String(text),
  options: {
    fontFace: 'Arial',
    fontSize: 10,
    color: TEXT,
    fill: { color: alt ? ROW_ALT : 'FFFFFF' },
    valign: 'middle',
    align,
  },
});

const tableRows = [
  [
    { text: 'Cible', options: { ...headerOpts, align: 'left' } },
    { text: 'Acquéreur', options: { ...headerOpts, align: 'left' } },
    { text: 'Géographie', options: { ...headerOpts, align: 'left' } },
    { text: 'EV (M$)', options: { ...headerOpts, align: 'right' } },
    { text: 'EBITDA (M$)', options: { ...headerOpts, align: 'right' } },
    { text: 'EV/EBITDA', options: { ...headerOpts, align: 'right' } },
    { text: 'EV/Rev', options: { ...headerOpts, align: 'right' } },
  ],
  ...deals.map((d, i) => [
    cell(d.cible, i % 2 === 1),
    cell(d.acquereur, i % 2 === 1),
    cell(d.geo, i % 2 === 1),
    cell(fmtM(d.ev), i % 2 === 1, 'right'),
    cell(fmtM(d.ebitda), i % 2 === 1, 'right'),
    cell(`${fmtM(d.evEbitda)}x`, i % 2 === 1, 'right'),
    cell(`${fmtM(d.evRev)}x`, i % 2 === 1, 'right'),
  ]),
];

s3.addTable(tableRows, {
  x: 0.5,
  y: 1.2,
  w: 12.3,
  colW: [2.2, 2.3, 1.35, 1.2, 1.3, 1.35, 1.2],
  border: { type: 'solid', color: 'E2E8F0', pt: 0.75 },
  fontSize: 10,
});

s3.addText(
  `Médianes (échantillon) : EV/EBITDA ${fmtM(medianEvEbitda)}x · EV/Revenus ${fmtM(medianEvRev)}x · Données arrondies selon sources [à compléter].`,
  {
    x: 0.5,
    y: 3.55,
    w: 12.3,
    h: 0.55,
    fontSize: 10,
    color: MUTED,
    italic: true,
    fontFace: 'Arial',
  }
);

// --- Slide 4 : Analyse des multiples ---
const s4 = pptx.addSlide();
s4.addText('Analyse des Multiples', {
  x: 0.5,
  y: 0.35,
  w: 12,
  h: 0.65,
  fontSize: 26,
  bold: true,
  color: PRIMARY,
  fontFace: 'Arial',
});
addDividerLine(s4, 1.0);

s4.addText(
  [
    { text: '1) Dispersion extrême des EV/EBITDA\n', options: { bold: true, fontSize: 14, color: PRIMARY } },
    {
      text:
        `Les multiples observés s’étendent d’environ ${fmtM(20.2)}x (Infomedia / TPG) à près de ${fmtM(297.8)}x (Perfect Corp / Cyberlink). Un multiple EBITDA élevé peut mécaniquement résulter d’un dénominateur proche de zéro, ce qui en fait un indicateur à manier avec des ajustements (EBITDA normalisé, one-offs, SBC [à compléter]).\n\n`,
    },
    { text: '2) Arbitrage croissance vs rentabilité\n', options: { bold: true, fontSize: 14, color: PRIMARY } },
    {
      text:
        'Les acquéreurs stratégiques peuvent capitaliser sur des synergies (rev share, data, distribution) et tolérer une rentabilité comptable plus faible à court terme, ce qui peut gonfler les multiples sur un EBITDA faible si la trajectoire de marge projetée est forte [à compléter]. À l’inverse, une cible plus mature et « cash generative » tend à ancrer la valorisation sur un EBITDA robuste.\n\n',
    },
    { text: '3) Acquéreur stratégique vs investisseur financier\n', options: { bold: true, fontSize: 14, color: PRIMARY } },
    {
      text:
        'Cyberlink et dotDigital illustrent des logiques d’intégration produit / géographie avec potentiel de synergie opérationnelle. À contrario, une prise de participation par un sponsor tel que TPG Capital LP (LBO / PE) privilégie souvent une diligence sur la qualité des cash-flows, le plan de désendettement et des multiples d’entrée plus « ancrés » sur l’EBITDA récurrent — cohérent avec un multiple EV/EBITDA sensiblement plus bas sur Infomedia.\n\n',
    },
    {
      text: 'Synthèse : privilégier en peer set des entreprises comparables sur le profil de revenus récurrents, la marge normalisée et le type d’acquéreur ; croiser EV/EBITDA avec EV/Revenus et méthodes DCF / comps sectoriels [à compléter].',
      options: { italic: true, fontSize: 12, color: MUTED },
    },
  ],
  {
    x: 0.55,
    y: 1.15,
    w: 12.2,
    h: 5.6,
    fontSize: 12,
    color: TEXT,
    fontFace: 'Arial',
    valign: 'top',
    lineSpacingMultiple: 1.12,
  }
);

s4.addText('Sources : même base que slide Executive Summary [à compléter].', {
  x: 0.5,
  y: 6.75,
  w: 12,
  h: 0.35,
  fontSize: 9,
  color: MUTED,
  italic: true,
  fontFace: 'Arial',
});

await pptx.writeFile({ fileName: outFile });
console.log(`OK: ${outFile}`);
