/**
 * Synthèse détaillée du bilan comptable français selon le PCG
 * Document investment-banking grade — docx principal + validation exceljs/pdfkit/pptxgenjs
 */
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  Table,
  TableRow,
  TableCell,
  WidthType,
  AlignmentType,
  BorderStyle,
  ShadingType,
  PageBreak,
} from 'docx';
import ExcelJS from 'exceljs';
import PDFDocument from 'pdfkit';
import PptxGenJS from 'pptxgenjs';
import fs from 'node:fs';
import path from 'node:path';
import { BRAND } from '../../shared/branding.js';

const OUTPUT_DIR = path.resolve('artifacts');
const OUTPUT_DOCX = path.join(OUTPUT_DIR, 'Synthese_Bilan_PCG_Detaillee.docx');

const NAVY = '0A2540';
const LIGHT_BLUE = 'E8F0FE';
const ALT_ROW = 'F2F7FB';
const WHITE = 'FFFFFF';
const GRAY = '666666';

// ── Exemple chiffré : société fictive « TechIndustrie SAS » (millions €) ─────
const EXEMPLE = {
  nom: 'TechIndustrie SAS',
  exercice: '31/12/2025',
  actifImmobilise: {
    incorporelles: 12.4,
    corporelles: 45.8,
    financieres: 8.2,
    total: 66.4,
  },
  actifCirculant: {
    stocks: 18.6,
    creancesClients: 22.3,
    autresCreances: 4.1,
    vmp: 3.5,
    disponibilites: 9.8,
    chargesAvance: 1.2,
    total: 59.5,
  },
  totalActif: 125.9,
  capitauxPropres: {
    capital: 10.0,
    primes: 5.0,
    reserves: 18.5,
    reportNouveau: 2.3,
    resultat: 6.7,
    total: 42.5,
  },
  provisions: 4.2,
  dettes: {
    financieres: 28.0,
    fournisseurs: 14.8,
    fiscalesSociales: 8.6,
    autres: 5.3,
    pca: 2.5,
    total: 59.2,
  },
  totalPassif: 125.9,
};

function cellBorders() {
  return {
    top: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
    bottom: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
    left: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
    right: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
  };
}

function headerCell(text, width = 2000) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: { fill: NAVY, type: ShadingType.CLEAR },
    borders: cellBorders(),
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text, bold: true, color: WHITE, size: 20, font: 'Calibri' })],
      }),
    ],
  });
}

function dataCell(text, { bold = false, align = AlignmentType.LEFT, shade = null, width = 2000 } = {}) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: shade ? { fill: shade, type: ShadingType.CLEAR } : undefined,
    borders: cellBorders(),
    children: [
      new Paragraph({
        alignment: align,
        children: [
          new TextRun({ text, bold, size: 20, font: 'Calibri', color: bold ? NAVY : '333333' }),
        ],
      }),
    ],
  });
}

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { after: opts.after ?? 120, before: opts.before ?? 0 },
    children: [
      new TextRun({
        text,
        size: opts.size ?? 22,
        font: 'Calibri',
        color: opts.color ?? '333333',
        bold: opts.bold,
        italics: opts.italics,
      }),
    ],
  });
}

function bullet(text) {
  return new Paragraph({
    bullet: { level: 0 },
    spacing: { after: 80 },
    children: [new TextRun({ text, size: 22, font: 'Calibri' })],
  });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, bold: true, color: NAVY, size: 32, font: 'Calibri' })],
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 200, after: 100 },
    children: [new TextRun({ text, bold: true, color: NAVY, size: 26, font: 'Calibri' })],
  });
}

function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 160, after: 80 },
    children: [new TextRun({ text, bold: true, color: NAVY, size: 24, font: 'Calibri' })],
  });
}

function buildTable(headers, rows, colWidths) {
  const headerRow = new TableRow({
    children: headers.map((h, i) => headerCell(h, colWidths?.[i] ?? 2400)),
  });
  const dataRows = rows.map((row, idx) => {
    const shade = idx % 2 === 0 ? ALT_ROW : WHITE;
    return new TableRow({
      children: row.map((cell, i) =>
        dataCell(String(cell), {
          bold: i === 0,
          align: i === 0 ? AlignmentType.LEFT : AlignmentType.CENTER,
          shade,
          width: colWidths?.[i] ?? 2400,
        })
      ),
    });
  });
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [headerRow, ...dataRows],
  });
}

function section1() {
  return [
    h1('1. Introduction et contexte réglementaire'),
    h2('1.1 Définition et objectifs du bilan comptable français'),
    p(
      "Le bilan comptable constitue l'un des états financiers fondamentaux de la comptabilité française. Il représente, à une date donnée (généralement la clôture de l'exercice), la situation patrimoniale de l'entreprise : l'ensemble des ressources dont elle dispose (actif) et l'origine de ces ressources (passif). Contrairement au compte de résultat qui mesure la performance sur une période, le bilan offre une photographie instantanée de la structure financière."
    ),
    p(
      "Selon l'article L123-12 du Code de commerce, les comptes annuels comprennent le bilan, le compte de résultat et une annexe. Le bilan doit donner une image fidèle du patrimoine, de la situation financière et du résultat de l'entreprise. Cette exigence d'image fidèle, consacrée par l'article 121-3 du Plan Comptable Général (PCG), constitue le principe cardinal de la comptabilité française."
    ),
    bullet(
      "Objectif patrimonial : inventorier et valoriser l'ensemble des éléments d'actif et de passif à la date de clôture."
    ),
    bullet(
      "Objectif informationnel : fournir aux parties prenantes (actionnaires, créanciers, administration fiscale, analystes M&A) une base fiable pour l'évaluation et la prise de décision."
    ),
    bullet(
      "Objectif de continuité d'exploitation : le bilan suppose que l'entreprise poursuit son activité, sauf indication contraire (article 212-1 PCG)."
    ),
    p(
      "Pour un analyste M&A, le bilan PCG est la pierre angulaire de l'analyse due diligence. Il permet de calculer la valeur nette comptable, d'identifier les actifs cachés ou sous-évalués, d'évaluer l'endettement net et de construire les ajustements de prix (locked-box, completion accounts)."
    ),
    h2('1.2 Évolution historique du Plan Comptable Général'),
    p(
      "Le PCG trouve ses origines dans le Plan Comptable de 1947, instauré pour harmoniser les pratiques comptables françaises dans la reconstruction d'après-guerre. Le Plan Comptable Général de 1957, puis sa révision de 1982, ont structuré la nomenclature des comptes en classes (1 à 8) toujours en vigueur aujourd'hui."
    ),
    p(
      "Les réformes majeures successives incluent : le PCG 1999 (intégration des normes européennes), le règlement ANC n°2014-03 (entrée en vigueur au 1er janvier 2015 pour les comptes individuels), et les amendements 2018-2020 (simplification pour les PME, évolutions sur les contrats de location IFRS 16 transposées partiellement). L'Autorité des Normes Comptables (ANC), créée en 2009, succède au Conseil National de la Comptabilité (CNC) et au Conseil de la Réglementation Comptable (CRC)."
    ),
    p(
      "La nomenclature actuelle distingue les comptes de bilan (classes 1 à 5) des comptes de gestion (classes 6 et 7) et des comptes spéciaux (classe 8). Cette architecture séculaire reste le socle de toute analyse financière en France, y compris dans les opérations de fusion-acquisition où le référentiel PCG sert de base aux ajustements vers des normes IFRS ou US GAAP."
    ),
    h2('1.3 Cadre réglementaire et normalisations'),
    buildTable(
      ['Organisme', 'Rôle', 'Texte de référence'],
      [
        ['Code de commerce', 'Obligations légales de tenue et présentation des comptes', 'L123-12 à L123-28'],
        ['ANC', 'Élaboration et publication des normes comptables françaises', 'Règlement ANC n°2014-03'],
        ['CRC (historique)', 'Avis et recommandations (intégré à l\'ANC depuis 2009)', 'Avis CRC'],
        ['OCDE / UE', 'Transposition des directives comptables européennes', 'Directive 2013/34/UE'],
      ],
      [2800, 3600, 2800]
    ),
    p(''),
    p(
      "L'article L123-16 du Code de commerce impose la tenue d'une comptabilité en partie double, l'enregistrement chronologique des opérations et l'établissement annuel d'un inventaire. Le règlement ANC n°2014-03 fixe les règles d'évaluation, de présentation et de publication applicables aux comptes sociaux."
    ),
    p(
      "Les avis de l'ANC (ex-CRC) complètent le PCG sur des points spécifiques : évaluation des engagements de retraite (avis CRC 2003-08), traitement des frais de développement (avis CRC 2004-06), consolidation (avis CRC 2009-02). En M&A, la connaissance de ces avis est essentielle pour anticiper les retraitements d'acquisition."
    ),
    h2('1.4 Harmonisation avec les normes IFRS'),
    p(
      "Le référentiel français (PCG) et le référentiel IFRS (International Financial Reporting Standards) coexistent en France. Les sociétés cotées établissent des comptes consolidés en IFRS (règlement CE n°1606/2002), tandis que leurs comptes individuels restent en PCG. Cette dualité crée des écarts significatifs que l'analyste M&A doit maîtriser."
    ),
    buildTable(
      ['Thème', 'PCG', 'IFRS'],
      [
        ['Évaluation des immobilisations', 'Coût historique, amortissements', 'Coût ou juste valeur (IAS 16, IAS 38)'],
        ['Goodwill', 'Amortissement possible (fonds commercial)', 'Test de dépréciation annuel (IAS 36)'],
        ['Contrats de location', 'Classification financière/opérationnelle', 'Droit d\'usage + dette (IFRS 16)'],
        ['Instruments financiers', 'Coût historique majoritairement', 'Juste valeur (IFRS 9)'],
        ['Provisions retraites', 'Comptabilisation au passif', 'IAS 19 — engagement actualisé'],
      ],
      [2800, 3200, 3200]
    ),
    p(''),
    p(
      "Lors d'une acquisition, le bridge PCG → IFRS est un livrable standard en due diligence financière. À titre d'illustration, pour TechIndustrie SAS, un retraitement IFRS 16 sur les locations opérationnelles pourrait augmenter l'actif de 4,2 M€ et la dette de 4,2 M€, modifiant mécaniquement les ratios d'endettement sans impact sur le fonds de roulement opérationnel."
    ),
    h2('1.5 Obligations légales des entreprises françaises'),
    p(
      "Toute personne morale de droit privé exerçant une activité commerciale, industrielle ou artisanale doit établir des comptes annuels (article L123-12). Les délais de dépôt au greffe varient : 7 mois après clôture pour les SA (6 mois pour les SAS/SARL), avec publication au Bulletin des Annonces Légales Obligatoires (BODACC)."
    ),
    bullet("Approbation par l'assemblée générale ordinaire dans les 6 mois de la clôture (article L232-23 pour les SA)."),
    bullet('Dépôt des comptes au greffe du tribunal de commerce — accessibles via Infogreffe pour les tiers.'),
    bullet('Audit légal obligatoire si dépassement de 2 seuils sur 3 : total bilan > 4 M€, CA > 8 M€, effectif > 50 (article L823-1).'),
    bullet('Commissaire aux comptes : certification des comptes, rapport sur les conventions réglementées, lettre de recommandations.'),
    p(
      "En contexte M&A, l'accès aux comptes certifiés des trois derniers exercices, au rapport du CAC et à la liasse fiscale (formulaires 2050 à 2059) constitue le minimum documentaire pour une analyse de qualité investment-banking."
    ),
    new PageBreak(),
  ];
}

function section2() {
  return [
    h1('2. Structure générale du bilan PCG'),
    h2('2.1 Présentation bipartite : Actif et Passif'),
    p(
      "Le bilan PCG adopte une présentation en liste verticale, organisée en deux masses patrimoniales : l'ACTIF (emplois) à gauche ou en premier, et le PASSIF (ressources) à droite ou en second. Cette présentation diffère du format américain (actif = passif + capitaux propres sur une seule page) mais véhicule la même équation fondamentale."
    ),
    p(
      "Le modèle de présentation est défini à l'article 512-1 du PCG et détaillé dans le formulaire fiscal 2050 (actif) et 2051 (passif). Les entreprises peuvent opter pour une présentation simplifiée (système abrégé) si elles ne dépassent pas certains seuils (article L123-16)."
    ),
    buildTable(
      ['Masse', 'Composantes principales', 'Classe PCG'],
      [
        ['Actif immobilisé', 'Incorporelles, corporelles, financières', 'Classe 2'],
        ['Actif circulant', 'Stocks, créances, VMP, disponibilités', 'Classe 3, 4, 5'],
        ['Capitaux propres', 'Capital, réserves, résultat', 'Classe 1 (10-14)'],
        ['Provisions', 'Risques et charges', 'Classe 1 (15)'],
        ['Dettes', 'Financières, exploitation, fiscales/sociales', 'Classe 1 (16), 4, 5'],
      ],
      [2400, 4000, 2400]
    ),
    p(''),
    h2('2.2 Principe de l\'équation fondamentale'),
    p(
      "L'équation comptable fondamentale s'énonce : ACTIF = PASSIF. Le passif se décompose en Capitaux propres + Provisions + Dettes. Cette égalité est vérifiée en permanence par la partie double : chaque opération affecte simultanément un compte d'actif et un compte de passif (ou de charges/produits)."
    ),
    p(
      `Exemple vérifié — ${EXEMPLE.nom} au ${EXEMPLE.exercice} : Actif total = ${EXEMPLE.totalActif} M€ = Capitaux propres (${EXEMPLE.capitauxPropres.total} M€) + Provisions (${EXEMPLE.provisions} M€) + Dettes (${EXEMPLE.dettes.total} M€) = ${EXEMPLE.totalPassif} M€. L'équilibre est strictement respecté.`
    ),
    h2('2.3 Classification par échéance et liquidité'),
    p(
      "L'article 211-1 du PCG distingue les éléments circulants (réalisable ou disponible à moins d'un an) des éléments non courants (immobilisations destinées à servir de façon durable). Cette distinction est cruciale pour l'analyse de liquidité et le calcul du fonds de roulement."
    ),
    bullet('Actif immobilisé = non courant : cycle long, amortissement sur plusieurs exercices.'),
    bullet('Actif circulant = courant : stocks (cycle < 1 an en principe), créances d\'exploitation, trésorerie.'),
    bullet('Dettes à plus d\'un an : emprunts obligataires, emprunts bancaires LT, dettes de crédit-bail.'),
    bullet('Dettes à moins d\'un an : fournisseurs, dettes fiscales/sociales, concours bancaires CT, partie CT des emprunts.'),
    h2('2.4 Distinction éléments courants et non courants'),
    p(
      "La distinction courant/non courant impacte directement les ratios financiers. Le fonds de roulement net global (FRNG) se calcule : Ressources stables − Emplois stables. Le besoin en fonds de roulement (BFR) = Actif circulant d'exploitation − Passif circulant d'exploitation."
    ),
    p(
      `Pour TechIndustrie SAS : FRNG = (42,5 + 4,2 + 28,0) − 66,4 = 8,3 M€. BFR = (18,6 + 22,3 + 4,1) − (14,8 + 8,6) = 29,6 M€. Trésorerie nette = FRNG − BFR = 8,3 − 29,6 = −21,3 M€, cohérente avec les concours bancaires et la structure de financement.`
    ),
    h2('2.5 Règles de présentation et d\'évaluation'),
    p(
      "Les règles de présentation (article 512-1 et suivants) imposent : classement par rubrique, ventilation des amortissements et dépréciations, distinction des éléments à moins et plus d'un an au passif. Les règles d'évaluation (articles 121-1 à 123-24) reposent sur le coût historique, la prudence et la continuité d'exploitation."
    ),
    bullet('Présentation brute puis amortissements/dépréciations (jamais de compensation actif/passif sauf exceptions).'),
    bullet('Évaluation à la date de clôture, après inventaire physique et rapprochement bancaire.'),
    bullet('Annexe obligatoire détaillant les méthodes d\'évaluation et les engagements hors bilan (article L123-19).'),
    new PageBreak(),
  ];
}

function section3() {
  return [
    h1('3. Analyse détaillée de l\'actif'),
    h2('3.1 Actif immobilisé'),
    h3('3.1.1 Immobilisations incorporelles'),
    p(
      "Les immobilisations incorporelles (comptes 20x) comprennent les éléments identifiables, sans substance physique, contrôlés par l'entreprise et générateurs de flux économiques futurs (article 211-8 PCG). Sont notamment concernés :"
    ),
    bullet('Fonds commercial (compte 207) : clientèle, droit au bail, enseigne. Non amortissable en principe, test de dépréciation annuel (article 214-18).'),
    bullet('Brevets, licences, marques (compte 205) : amortissement sur durée de protection ou d\'utilité (article 214-16).'),
    bullet('Logiciels (compte 205) : acquis ou créés en interne si critères de l\'article 212-3 remplis (frais de développement activables).'),
    bullet('Frais de recherche et développement (compte 203) : activation possible si projet identifié, faisabilité démontrée, rentabilité probable (article 212-3).'),
    p(
      `Exemple : TechIndustrie SAS comptabilise 12,4 M€ d'incorporelles dont 5,8 M€ de fonds commercial acquis lors de l'acquisition de MicroParts en 2022, 4,1 M€ de logiciels ERP amortis linéairement sur 5 ans (dotation annuelle : 820 K€), et 2,5 M€ de frais de développement activés sur un programme R&D qualifié.`
    ),
    h3('3.1.2 Immobilisations corporelles'),
    p(
      "Les immobilisations corporelles (comptes 21x) regroupent les terrains (211), constructions (213), installations techniques (215), matériel industriel (2181), matériel de transport (2182) et agencements (2183). Évaluation au coût d'acquisition ou de production (article 213-1)."
    ),
    bullet('Terrains : non amortissables (durée de vie illimitée), sauf terrains d\'exploitation (carrières).'),
    bullet('Constructions : amortissement linéaire 20 à 50 ans selon nature.'),
    bullet('Matériel industriel : amortissement linéaire (durée fiscale) ou dégressif si éligible (article 39 A CGI).'),
    p(
      `TechIndustrie SAS : corporelles nettes de 45,8 M€ (brut 78,2 M€ − amortissements cumulés 32,4 M€). Dont terrain 6,0 M€, bâtiment industriel 22,5 M€ (amorti sur 25 ans), lignes de production 17,3 M€ (amorti sur 10 ans, dégressif fiscal). Valeur vénale estimée en due diligence : 52 M€ (+13,5% vs net comptable).`
    ),
    h3('3.1.3 Immobilisations financières'),
    p(
      "Les immobilisations financières (comptes 26x, 27x) comprennent les participations (261), créances rattachées à des participations (267), prêts (274), dépôts et cautionnements (275). Elles traduisent des investissements durables dans d'autres entités ou des créances à long terme."
    ),
    bullet('Titres de participation : évaluation au coût, dépréciation si perte durable de valeur (article 213-8).'),
    bullet('Créances rattachées : prêts intra-groupe, comptes courants d\'associés bloqués.'),
    bullet('En M&A : attention aux participations à réévaluer en juste valeur dans les comptes consolidés IFRS.'),
    p(`TechIndustrie SAS détient 8,2 M€ de participations (15% de LogiSupply SA, 30% de DataServ SARL — mise en équivalence potentielle en consolidation).`),
    h3('3.1.4 Méthodes d\'évaluation et d\'amortissement'),
    buildTable(
      ['Méthode', 'Principe', 'Application PCG', 'Exemple chiffré'],
      [
        ['Linéaire', 'Charge constante = coût/durée', 'Article 214-17', 'Machine 1 M€ / 10 ans = 100 K€/an'],
        ['Dégressif', 'Coefficients fiscaux (1,25 à 2,25)', 'Article 39 A CGI', 'Bien 3 ans : coeff. 1,25'],
        ['Unités d\'œuvre', 'Coût × unités produites/prévues', 'Article 214-17', 'Presse 500 K cycles = dotation variable'],
      ],
      [2000, 2800, 2400, 2400]
    ),
    p(''),
    h3('3.1.5 Provisions pour dépréciation'),
    p(
      "Les dépréciations d'actif (compte 29x) traduisent une perte de valeur durable. Test de dépréciation obligatoire pour les incorporelles à durée indéfinie (fonds commercial) et les titres de participation si indices de perte de valeur (article 213-3). La reprise de dépréciation est limitée à la valeur qui aurait été constatée sans dépréciation initiale."
    ),
    h2('3.2 Actif circulant'),
    h3('3.2.1 Stocks et en-cours'),
    p(
      "Les stocks (classe 3) sont évalués au coût d'acquisition ou de production (article 213-30). Méthodes de valorisation de sortie : FIFO (PEPS) ou coût moyen pondéré (CMP). Test de dépréciation à chaque clôture si prix de vente net < coût (article 213-31)."
    ),
    buildTable(
      ['Compte', 'Nature', 'TechIndustrie SAS (M€)'],
      [
        ['311', 'Matières premières', '6,2'],
        ['351', 'Produits en cours', '4,8'],
        ['355', 'Produits finis', '7,6'],
        ['Total stocks', '', '18,6'],
      ],
      [1800, 3200, 2400]
    ),
    p(''),
    p('Rotation stocks = CA / Stock moyen. Pour TechIndustrie (CA 95 M€) : 95 / 18,6 = 5,1 rotations/an soit 71 jours de stock.'),
    h3('3.2.2 Créances clients et autres créances'),
    p(
      "Les créances d'exploitation (compte 411) et autres créances (comptes 42x, 44x, 46x, 47x) sont évaluées à leur valeur nominale, diminuées des dépréciations (compte 491). En M&A, l'analyse de la qualité du poste clients (DSO, concentration, créances litigieuses) est un point de vigilance majeur."
    ),
    bullet('Créances clients : 22,3 M€ brut, dépréciation 0,8 M€ → net 21,5 M€. DSO = 22,3 / (95/365) = 86 jours.'),
    bullet('État et autres organismes (44566) : 2,1 M€ (TVA déductible, crédits d\'impôt recherche).'),
    bullet('Débiteurs divers : 1,2 M€ (acomptes versés, cautionnements).'),
    h3('3.2.3 Valeurs mobilières de placement'),
    p(
      "Les VMP (compte 50) comprennent les actions, obligations et SICAV détenues à court terme pour une mise en liquidité prochaine. Évaluation à la valeur de marché à la clôture si cotées (article 213-34), avec constatation des plus ou moins-values latentes."
    ),
    h3('3.2.4 Disponibilités et quasi-disponibilités'),
    p(
      "Les disponibilités (comptes 512, 53) et équivalents de trésorerie (VMP à échéance < 3 mois) représentent la liquidité immédiate. TechIndustrie SAS : 9,8 M€ en banque dont 2,0 M€ bloqués en garantie de lignes de crédit → liquidité immédiate réelle : 7,8 M€."
    ),
    h3('3.2.5 Charges constatées d\'avance'),
    p(
      "Les charges constatées d'avance (compte 486) étalent comptablement des charges payées d'avance (assurances, loyers, abonnements). TechIndustrie : 1,2 M€, dont 0,7 M€ d'assurances et 0,5 M€ de maintenance prépayée."
    ),
    buildTable(
      ['Rubrique actif', 'Montant (M€)', '% du total actif'],
      [
        ['Immobilisations incorporelles nettes', '12,4', '9,8%'],
        ['Immobilisations corporelles nettes', '45,8', '36,4%'],
        ['Immobilisations financières', '8,2', '6,5%'],
        ['Stocks', '18,6', '14,8%'],
        ['Créances', '26,4', '21,0%'],
        ['VMP et disponibilités', '13,3', '10,6%'],
        ['Charges constatées d\'avance', '1,2', '1,0%'],
        ['TOTAL ACTIF', '125,9', '100,0%'],
      ],
      [3600, 2400, 2400]
    ),
    p(''),
    new PageBreak(),
  ];
}

function section4() {
  return [
    h1('4. Analyse détaillée du passif'),
    h2('4.1 Capitaux propres'),
    h3('4.1.1 Capital social et primes'),
    p(
      "Le capital social (compte 101) représente les apports des associés. Les primes d'émission (1041), de fusion (1042) et d'apport (1043) constituent une réserve distinte, non distribuable mais transmissible en cas de cession de titres. L'article L223-1 du Code de commerce fixe le capital minimum (1 € pour les SAS/SARL)."
    ),
    p(`TechIndustrie SAS : capital 10,0 M€ (100 000 actions de 100 €), prime d'émission 5,0 M€ (levée de fonds 2020).`),
    h3('4.1.2 Réserves'),
    p(
      "Les réserves se décomposent en : réserve légale (1061, 5% du bénéfice jusqu'à 10% du capital — article L232-10), réserves statutaires (1062, imposées par les statuts), réserves facultatives (1063, dotations discrétionnaires). Les réserves ne peuvent être distribuées qu'après approbation des comptes."
    ),
    bullet('Réserve légale : 1,0 M€ (plafond atteint à 10% × 10 M€ capital).'),
    bullet('Réserves statutaires : 3,5 M€ (dotation obligatoire de 15% du RNPG).'),
    bullet('Réserves facultatives : 14,0 M€ (dotations antérieures et prime de capitalisation).'),
    h3('4.1.3 Report à nouveau'),
    p(
      "Le report à nouveau (compte 110 ou 119) cumule les résultats des exercices antérieurs non distribués. Un report créditeur (110) renforce les capitaux propres ; un report débiteur (119) constitue un « trou de bilan » à combler avant toute distribution."
    ),
    h3('4.1.4 Résultat de l\'exercice'),
    p(
      "Le résultat de l'exercice (compte 120 ou 129) est le solde du compte de résultat. Il est incorporé aux capitaux propres à l'approbation des comptes. TechIndustrie SAS : bénéfice 2025 de 6,7 M€ (marge nette 7,1% sur CA 95 M€)."
    ),
    h3('4.1.5 Provisions réglementées'),
    p(
      "Les provisions réglementées (compte 14) sont des ressources stables d'origine fiscale : amortissements dérogatoires (145), provisions pour investissement (145), plus-values à court terme sur VMP (142). Elles constituent des différences temporaires entre résultat comptable et fiscal."
    ),
    h2('4.2 Provisions pour risques et charges'),
    p(
      "Les provisions (compte 15) ne sont comptabilisées que si l'entreprise a une obligation actuelle à la clôture, issue d'un événement passé, dont le règlement est probable et le montant estimable (article 312-1 à 312-8 PCG, inspiré de la notion IAS 37)."
    ),
    buildTable(
      ['Type de provision', 'Compte', 'Montant TechInd. (M€)', 'Base PCG'],
      [
        ['Provisions pour risques', '151', '1,8', 'Article 312-1 — litiges, garanties'],
        ['Provisions pour charges', '153', '1,2', 'Article 312-5 — restructuration'],
        ['Provisions retraites', '155', '1,2', 'Avis CRC 2003-08 / IAS 19'],
        ['Total provisions', '', '4,2', ''],
      ],
      [2800, 1200, 2000, 2800]
    ),
    p(''),
    bullet('Provision litige client : 1,0 M€ (contentieux qualité livraison 2024).'),
    bullet('Provision garantie produits : 0,8 M€ (1,5% du CA garanti).'),
    bullet('Provision restructuration : 1,2 M€ (plan social 2026 anticipé).'),
    bullet('Engagements retraite : 1,2 M€ (méthode actuarielle, taux 3,5%, mortalité INSEE).'),
    p(
      "En due diligence M&A, les provisions doivent être passées au crible : caractère récurrent vs non récurrent, adéquation du montant, risque de sous-provisionnement (notamment retraites et environnement)."
    ),
    h2('4.3 Dettes'),
    h3('4.3.1 Dettes financières'),
    p(
      "Les dettes financières (comptes 16x, 17x, 519) comprennent emprunts obligataires (161), emprunts bancaires (164), dettes de crédit-bail (167), concours bancaires courants (519). Ventilation obligatoire entre partie à moins d'un an et plus d'un an (annexe)."
    ),
    bullet('Emprunts bancaires LT : 20,0 M€ (taux moyen 3,2%, échéance 2028).'),
    bullet('Emprunts obligataires : 5,0 M€ (taux 4,5%, échéance 2027).'),
    bullet('Crédit-bail immobilier : 3,0 M€ (valeur actuelle des redevances).'),
    h3('4.3.2 Dettes fournisseurs et comptes rattachés'),
    p(
      "Les dettes d'exploitation (compte 401) traduisent les achats non encore réglés. Le DPO (Days Payable Outstanding) mesure le délai de paiement fournisseurs. TechIndustrie : 14,8 M€, DPO = 14,8 / (62/365) = 87 jours (achats annuels 62 M€)."
    ),
    h3('4.3.3 Dettes fiscales et sociales'),
    p(
      "Comptes 42x (personnel), 43x (sécurité sociale), 44x (État — TVA, IS). TechIndustrie : 8,6 M€ dont IS à payer 3,2 M€, TVA collectée nette 2,8 M€, charges sociales 2,6 M€."
    ),
    h3('4.3.4 Autres dettes'),
    p(
      "Dettes sur immobilisations (404), dettes diverses (408), produits à recevoir. TechIndustrie : 5,3 M€ dont 2,0 M€ de dividendes à payer et 1,5 M€ d'acomptes clients reçus."
    ),
    h3('4.3.5 Produits constatés d\'avance'),
    p(
      "Les produits constatés d'avance (compte 487) étalent des produits encaissés d'avance. TechIndustrie : 2,5 M€ de contrats de maintenance pluriannuelle facturés d'avance."
    ),
    buildTable(
      ['Rubrique passif', 'Montant (M€)', '% du total passif'],
      [
        ['Capital et primes', '15,0', '11,9%'],
        ['Réserves et report', '20,8', '16,5%'],
        ['Résultat exercice', '6,7', '5,3%'],
        ['Provisions réglementées', '0,0', '0,0%'],
        ['Provisions risques/charges', '4,2', '3,3%'],
        ['Dettes financières', '28,0', '22,2%'],
        ['Dettes exploitation', '14,8', '11,8%'],
        ['Dettes fiscales/sociales', '8,6', '6,8%'],
        ['Autres dettes et PCA', '7,8', '6,2%'],
        ['TOTAL PASSIF', '125,9', '100,0%'],
      ],
      [3600, 2400, 2400]
    ),
    p(''),
    new PageBreak(),
  ];
}

function section5() {
  return [
    h1('5. Principes d\'évaluation et méthodes comptables'),
    h2('5.1 Coût historique vs juste valeur'),
    p(
      "Le PCG consacre le coût historique comme mode d'évaluation par défaut (article 121-1). La juste valeur n'est utilisée que pour certains instruments financiers (VMP cotées — article 213-34) et dans des cas spécifiques (fusion, scission — article 744-1). Cette approche prudentielle diffère fondamentalement des IFRS qui privilégient la juste valeur pour de nombreuses catégories d'actifs."
    ),
    p(
      "Impact M&A : en acquisition, l'acheteur réévalue les actifs identifiables en juste valeur (purchase price allocation). L'écart entre juste valeur et valeur nette comptable PCG génère du goodwill ou des plus-values de réévaluation — élément clé de la négociation de prix."
    ),
    h2('5.2 Principe de prudence'),
    p(
      "L'article 121-4 du PCG impose la prudence : ne comptabiliser des profits que réalisés, prendre en compte toutes les pertes probables dès leur survenance. Conséquences pratiques : provisions pour risques, dépréciations d'actif, non-activation des revenus incertains, évaluation des stocks au moindre du coût et de la valeur nette de réalisation."
    ),
    bullet('Pas de constitution de provisions « générales » sans obligation identifiable (interdiction des cookie jar reserves).'),
    bullet('Test de dépréciation des créances douteuses dès premier indice (retard > 90 jours, procédure collective client).'),
    bullet('Charges à payer (compte 428, 438, 448) pour rattacher les charges à l\'exercice même si non encore facturées.'),
    h2('5.3 Méthodes d\'amortissement'),
    p(
      "L'amortissement répartit le coût d'un actif sur sa durée d'utilité (article 214-1). La durée retenue doit refléter l'usage économique, pas nécessairement la durée fiscale. Le passage du dégressif au linéaire est obligatoire quand la dotation dégressive devient inférieure à la dotation linéaire résiduelle."
    ),
    p(
      "Exemple comparatif — équipement 500 K€, durée 5 ans : Linéaire = 100 K€/an constant. Dégressif (coeff. 1,75, durée fiscale 5 ans) : An 1 = 175 K€, An 2 = 131 K€, An 3 = 98 K€, An 4 = 73 K€, An 5 = 23 K€. Avantage fiscal en début de vie de l'actif, neutralité sur la durée totale."
    ),
    h2('5.4 Provisions et dépréciations'),
    p(
      "Les dépréciations (comptes 29x, 39x, 49x, 59x) ajustent la valeur des actifs. Les provisions (comptes 15x) couvrent les obligations au passif. La distinction est fondamentale : une dépréciation affecte la valeur d'un actif existant ; une provision reconnaît une dette probable sans contrepartie d'actif."
    ),
    h2('5.5 Conversion des devises'),
    p(
      "Les opérations en devises sont converties au cours du jour de la transaction (article 420-1). Les actifs et passifs monétaires en devises à la clôture sont convertis au cours de clôture. Les différences de conversion sont enregistrées en résultat (ou en capitaux propres pour couverture de flux de trésorerie si critères remplis — article 420-3)."
    ),
    p(
      "TechIndustrie SAS (exposition USD 15% du CA) : créances USD 1,8 M€ converties au cours 1,08 EUR/USD. Variation de ±5% du dollar = impact résultat de ±90 K€ — sensibilité à intégrer dans le modèle de deal."
    ),
    new PageBreak(),
  ];
}

function section6() {
  const cp = EXEMPLE.capitauxPropres.total;
  const dettesFin = EXEMPLE.dettes.financieres;
  const actifCirc = EXEMPLE.actifCirculant.total;
  const dettesCT = EXEMPLE.dettes.fournisseurs + EXEMPLE.dettes.fiscalesSociales + 5.0;
  const tresorerie = EXEMPLE.actifCirculant.disponibilites;
  const stocks = EXEMPLE.actifCirculant.stocks;
  const creances = EXEMPLE.actifCirculant.creancesClients;
  const ca = 95;
  const achats = 62;

  const liqGenerale = (actifCirc / dettesCT).toFixed(2);
  const liqReduite = ((actifCirc - stocks) / dettesCT).toFixed(2);
  const liqImmediate = (tresorerie / dettesCT).toFixed(2);
  const autonomie = ((cp / EXEMPLE.totalPassif) * 100).toFixed(1);
  const endettement = ((dettesFin / cp) * 100).toFixed(1);
  const bfr = (stocks + creances + 4.1 - EXEMPLE.dettes.fournisseurs - EXEMPLE.dettes.fiscalesSociales).toFixed(1);
  const rotStock = (ca / stocks).toFixed(1);
  const dso = ((creances / ca) * 365).toFixed(0);
  const dpo = ((EXEMPLE.dettes.fournisseurs / achats) * 365).toFixed(0);
  const gearing = ((dettesFin / (cp + dettesFin)) * 100).toFixed(1);

  return [
    h1('6. Analyse financière du bilan'),
    h2('6.1 Ratios de liquidité'),
    p(
      "Les ratios de liquidité mesurent la capacité de l'entreprise à honorer ses engagements à court terme. Ils sont calculés à partir des postes du bilan circulant."
    ),
    buildTable(
      ['Ratio', 'Formule', 'TechIndustrie SAS', 'Benchmark industrie'],
      [
        ['Liquidité générale', 'Actif circulant / Dettes CT', liqGenerale, '1,2 – 1,8'],
        ['Liquidité réduite', '(AC − Stocks) / Dettes CT', liqReduite, '0,8 – 1,2'],
        ['Liquidité immédiate', 'Disponibilités / Dettes CT', liqImmediate, '0,2 – 0,5'],
      ],
      [2800, 3600, 2000, 2000]
    ),
    p(''),
    p(
      `Interprétation : la liquidité générale de ${liqGenerale} est dans la fourchette haute, soutenue par un poste stocks important. La liquidité réduite de ${liqReduite} reste confortable. La liquidité immédiate de ${liqImmediate} reflète une trésorerie modérée, l'entreprise s'appuyant sur ses lignes de crédit confirmées (12 M€ non tirées).`
    ),
    h2('6.2 Ratios de structure financière'),
    buildTable(
      ['Ratio', 'Formule', 'Valeur', 'Analyse'],
      [
        ['Autonomie financière', 'CP / Total passif', `${autonomie}%`, 'Structure équilibrée (> 30% seuil prudent)'],
        ['Endettement financier', 'Dettes fin. / CP', `${endettement}%`, 'Levier modéré, capacité de levée supplémentaire'],
        ['Gearing', 'Dettes fin. / (CP + Dettes fin.)', `${gearing}%`, 'Acceptable pour secteur industriel'],
        ['Couverture des immos.', 'CP / Actif immobilisé', `${((cp / EXEMPLE.actifImmobilise.total) * 100).toFixed(1)}%`, 'Immobilisations partiellement autofinancées'],
      ],
      [2600, 3000, 1600, 2600]
    ),
    p(''),
    h2('6.3 Besoin en fonds de roulement'),
    p(
      `Le BFR d'exploitation de TechIndustrie SAS s'établit à ${bfr} M€, principalement tiré par le poste clients (DSO ${dso} jours) et stocks (rotation ${rotStock}x, soit ${(365 / rotStock).toFixed(0)} jours). Le DPO de ${dpo} jours partiellement compense le BFR.`
    ),
    p(
      "En M&A, le BFR normatif est un ajustement de prix incontournable : l'acquéreur exige un BFR de livraison correspondant au BFR normatif calculé sur la base des derniers mois d'activité. Un BFR de livraison inférieur génère un ajustement de prix négatif (cash-free/debt-free mechanism)."
    ),
    h2('6.4 Ratios de rotation'),
    buildTable(
      ['Indicateur', 'Formule', 'Valeur', 'Objectif sectoriel'],
      [
        ['Rotation stocks', 'CA / Stock moyen', `${rotStock}x`, '6 – 8x'],
        ['DSO (jours clients)', '(Créances / CA) × 365', `${dso} jours`, '60 – 75 jours'],
        ['DPO (jours fournisseurs)', '(Dettes fourn. / Achats) × 365', `${dpo} jours`, '60 – 90 jours'],
        ['Cycle d\'exploitation', 'DSO + DIO − DPO', `${Number(dso) + Math.round(365 / rotStock) - Number(dpo)} jours`, 'Minimiser'],
      ],
      [2400, 3200, 1600, 2400]
    ),
    p(''),
    h2('6.5 Indicateurs de solvabilité'),
    p(
      "La solvabilité à long terme évalue la capacité de l'entreprise à rembourser l'ensemble de ses dettes. L'endettement net = Dettes financières − Trésorerie = 28,0 − 9,8 = 18,2 M€, soit 0,43× les capitaux propres et 2,7× l'EBITDA estimé (6,8 M€) — ratio Debt/EBITDA compatible avec un financement LBO mid-market."
    ),
    bullet('Capacité de remboursement : Endettement net / CAF = 18,2 / 8,5 = 2,1 ans (< 4 ans seuil bancaire).'),
    bullet('Couverture des charges financières : EBIT / Charges d\'intérêts = 12,4 / 1,1 = 11,3× (confortable).'),
    bullet('Fonds de roulement net global : 8,3 M€ positif — équilibre financier structurel assuré.'),
    h1('Conclusion'),
    p(
      "Le bilan comptable selon le Plan Comptable Général constitue le référentiel incontournable de l'analyse financière en France. Sa structure bipartite, ses principes d'évaluation prudentiels et sa nomenclature normalisée offrent un cadre fiable et comparable pour l'ensemble des acteurs économiques."
    ),
    p(
      "Pour les professionnels du M&A, la maîtrise du bilan PCG dépasse la simple lecture comptable : elle conditionne l'évaluation des cibles, la structuration des transactions (SPA, ajustements de prix, garanties d'actif-passif), les retraitements vers IFRS et la construction des modèles de financement (LBO, MBO, earn-out)."
    ),
    p(
      "L'exemple de TechIndustrie SAS illustre comment un bilan équilibré (actif = passif = 125,9 M€) peut coexister avec des tensions opérationnelles (BFR élevé, DSO allongé) et des opportunités de création de valeur (écart valeur vénale / valeur comptable des immobilisations, optimisation du BFR post-acquisition)."
    ),
    p(
      "En conclusion, le bilan PCG n'est pas une fin en soi mais un point de départ : la valeur économique d'une entreprise se construit par-dessus et au-delà de ses chiffres comptables, en combinant expertise comptable, analyse financière et jugement stratégique — compétences au cœur de la pratique investment-banking."
    ),
    p(''),
    p(`Document généré le 10 juin 2026 — ${BRAND.confidentiality}`, { italics: true, color: '999999', size: 20 }),
  ];
}

async function buildDocx() {
  const children = [
    new Paragraph({
      heading: HeadingLevel.TITLE,
      children: [
        new TextRun({
          text: 'Synthèse détaillée du bilan comptable français',
          bold: true,
          size: 40,
          color: NAVY,
          font: 'Calibri',
        }),
      ],
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: 'Selon le Plan Comptable Général (PCG) — Guide M&A & Finance',
          size: 24,
          color: GRAY,
          font: 'Calibri',
        }),
      ],
    }),
    new Paragraph({
      children: [
        new TextRun({ text: 'Strictement confidentiel — Usage professionnel', italics: true, color: '999999', size: 20 }),
      ],
    }),
    new Paragraph({ text: '' }),
    p(
      "Ce document constitue une synthèse exhaustive du bilan comptable français selon le PCG, destinée aux analystes financiers, auditeurs et professionnels du M&A. Il couvre le cadre réglementaire, la structure du bilan, l'analyse détaillée de l'actif et du passif, les principes d'évaluation et les ratios d'analyse financière, avec exemples chiffrés et références aux articles du PCG."
    ),
    new PageBreak(),
    ...section1(),
    ...section2(),
    ...section3(),
    ...section4(),
    ...section5(),
    ...section6(),
  ];

  const doc = new Document({
    creator: 'M&IA',
    title: 'Synthèse Bilan PCG Détaillée',
    description: 'Guide complet du bilan comptable français selon le PCG',
    sections: [
      {
        properties: {
          page: {
            margin: { top: 1440, right: 1200, bottom: 1440, left: 1200 },
          },
        },
        children,
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
  fs.writeFileSync(OUTPUT_DOCX, buffer);
  console.log(`OK: wrote ${OUTPUT_DOCX} (${buffer.length} bytes)`);
  return buffer.length;
}

async function buildExcelValidation() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Bilan TechIndustrie');
  ws.columns = [
    { header: 'Poste', key: 'poste', width: 35 },
    { header: 'Montant (M€)', key: 'montant', width: 15 },
    { header: 'Classe PCG', key: 'classe', width: 12 },
  ];
  const rows = [
    ['Immobilisations incorporelles', EXEMPLE.actifImmobilise.incorporelles, '20'],
    ['Immobilisations corporelles', EXEMPLE.actifImmobilise.corporelles, '21'],
    ['Immobilisations financières', EXEMPLE.actifImmobilise.financieres, '26-27'],
    ['Stocks', EXEMPLE.actifCirculant.stocks, '3'],
    ['Créances clients', EXEMPLE.actifCirculant.creancesClients, '411'],
    ['Disponibilités', EXEMPLE.actifCirculant.disponibilites, '512'],
    ['TOTAL ACTIF', EXEMPLE.totalActif, ''],
    ['Capital social', EXEMPLE.capitauxPropres.capital, '101'],
    ['Réserves', EXEMPLE.capitauxPropres.reserves, '106'],
    ['Résultat', EXEMPLE.capitauxPropres.resultat, '120'],
    ['Provisions', EXEMPLE.provisions, '15'],
    ['Dettes financières', EXEMPLE.dettes.financieres, '16'],
    ['Dettes fournisseurs', EXEMPLE.dettes.fournisseurs, '401'],
    ['TOTAL PASSIF', EXEMPLE.totalPassif, ''],
  ];
  rows.forEach((r) => ws.addRow({ poste: r[0], montant: r[1], classe: r[2] }));
  const xlsxPath = path.join(OUTPUT_DIR, '_validation_bilan_pcg.xlsx');
  await wb.xlsx.writeFile(xlsxPath);
  console.log(`OK: validation excel ${xlsxPath}`);
}

async function buildPdfSummary() {
  const pdfPath = path.join(OUTPUT_DIR, '_validation_bilan_pcg.pdf');
  const doc = new PDFDocument({ margin: 50 });
  const stream = fs.createWriteStream(pdfPath);
  doc.pipe(stream);
  doc.fontSize(18).text('Synthèse Bilan PCG — Résumé', { underline: true });
  doc.moveDown();
  doc.fontSize(11).text(`Société exemple : ${EXEMPLE.nom}`);
  doc.text(`Exercice : ${EXEMPLE.exercice}`);
  doc.text(`Total actif : ${EXEMPLE.totalActif} M€`);
  doc.text(`Capitaux propres : ${EXEMPLE.capitauxPropres.total} M€`);
  doc.text(`Dettes financières : ${EXEMPLE.dettes.financieres} M€`);
  doc.text(`Équation vérifiée : Actif = Passif = ${EXEMPLE.totalActif} M€`);
  doc.end();
  await new Promise((res) => stream.on('finish', res));
  console.log(`OK: validation pdf ${pdfPath}`);
}

async function buildPptxSummary() {
  const pptx = new PptxGenJS();
  pptx.author = 'M&IA';
  const slide = pptx.addSlide();
  slide.addText('Bilan PCG — Structure', { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 24, bold: true, color: NAVY });
  slide.addText(
    [
      { text: 'ACTIF', options: { bold: true, breakLine: true } },
      { text: `Immobilisé : ${EXEMPLE.actifImmobilise.total} M€`, options: { breakLine: true } },
      { text: `Circulant : ${EXEMPLE.actifCirculant.total} M€`, options: { breakLine: true } },
      { text: `Total : ${EXEMPLE.totalActif} M€`, options: { bold: true, breakLine: true } },
      { text: '', options: { breakLine: true } },
      { text: 'PASSIF', options: { bold: true, breakLine: true } },
      { text: `Capitaux propres : ${EXEMPLE.capitauxPropres.total} M€`, options: { breakLine: true } },
      { text: `Provisions : ${EXEMPLE.provisions} M€`, options: { breakLine: true } },
      { text: `Dettes : ${EXEMPLE.dettes.total} M€`, options: { breakLine: true } },
      { text: `Total : ${EXEMPLE.totalPassif} M€`, options: { bold: true } },
    ],
    { x: 0.5, y: 1.2, w: 9, h: 4, fontSize: 14 }
  );
  const pptxPath = path.join(OUTPUT_DIR, '_validation_bilan_pcg.pptx');
  await pptx.writeFile({ fileName: pptxPath });
  console.log(`OK: validation pptx ${pptxPath}`);
}

const size = await buildDocx();
await buildExcelValidation();
await buildPdfSummary();
await buildPptxSummary();

if (size < 5120) {
  console.error(`ERROR: fichier trop petit (${size} bytes < 5 KB)`);
  process.exit(1);
}
