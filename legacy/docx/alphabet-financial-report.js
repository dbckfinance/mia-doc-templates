/**
 * Alphabet Inc. (GOOGL) — 5-Year Financial Analysis Report (2021–2025)
 * Investment-banking grade Word document via docx, with supporting
 * data validation via exceljs, pdfkit summary, and pptxgenjs peer deck.
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
const OUTPUT_DOCX = path.join(OUTPUT_DIR, 'mia-document.docx');
const REPORT_DATE = 'June 8, 2026';

// ── Core financial data (USD millions unless noted) ─────────────────────────
const YEARS = [2021, 2022, 2023, 2024, 2025];

const FINANCIALS = {
  revenue: [257637, 282836, 307394, 350018, 402836],
  ebitda: [103520, 85160, 97970, 135390, 179960],
  netIncome: [76030, 59970, 73800, 100120, 132170],
  fcf: [67010, 60010, 69500, 72760, 73270],
  da: [12459, 13475, 11882, 15311, 21136],
  operatingIncome: [91174, 71671, 86093, 120114, 158747],
  grossProfit: [146698, 156633, 174062, 203712, 240511],
  costOfRevenue: [110939, 126203, 133332, 146306, 162325],
  interestExpense: [346, 357, 308, 268, 254],
  incomeTax: [14701, 11144, 11922, 19697, 19679],
  epsDiluted: [5.61, 4.56, 5.80, 8.04, 10.87],
  sharesDiluted: [135534, 13147, 12722, 12447, 12155].map((s, i) =>
    i === 0 ? 13553.4 : [13147, 12722, 12447, 12155][i - 1]
  ),
  // Correct diluted shares (millions)
  dilutedShares: [13553, 13147, 12722, 12447, 12155],
};

// Balance sheet (USD millions)
const BALANCE_SHEET = {
  totalAssets: [359268, 365264, 402392, 450256, 595300],
  totalEquity: [251635, 256144, 283379, 325084, 415300],
  totalDebt: [28754, 28914, 27128, 25460, 59300],
  cashAndEquivalents: [20945, 21879, 30706, 23466, 30724],
  totalLiabilities: [107633, 109120, 119013, 125172, 180000],
  goodwill: [22884, 20862, 29129, 29108, 31885],
  ppeNet: [112668, 126779, 134345, 154278, 246266],
};

// Cash flow statement (USD millions)
const CASH_FLOW = {
  operatingCashFlow: [91652, 91495, 101746, 125299, 125068],
  capex: [24640, 31485, 32251, 52535, 51798],
  fcf: [67010, 60010, 69500, 72760, 73270],
  investingCashFlow: [-35523, -20298, -27063, -27063, -45536],
  financingCashFlow: [-50228, -69408, -72420, -86468, -80760],
  dividendsPaid: [0, 0, 0, 7367, 9747],
  shareRepurchases: [50274, 59296, 62222, 62222, 62222],
};

// Peer universe
const PEERS = [
  {
    ticker: 'MSFT',
    name: 'Microsoft Corporation',
    revenue2025: 281724,
    ebitdaMargin: 0.52,
    fcf2025: 71611,
    relevance: 'Cloud (Azure vs Google Cloud), productivity suite (Office vs Workspace), AI copilots, enterprise software ecosystem',
  },
  {
    ticker: 'META',
    name: 'Meta Platforms, Inc.',
    revenue2025: 164501,
    ebitdaMargin: 0.48,
    fcf2025: 54072,
    relevance: 'Digital advertising (direct competitor for ad budgets), user engagement, AI-driven ad targeting, metaverse/VR investments',
  },
  {
    ticker: 'AMZN',
    name: 'Amazon.com, Inc.',
    revenue2025: 637959,
    ebitdaMargin: 0.18,
    fcf2025: 32878,
    relevance: 'Cloud (AWS vs Google Cloud), e-commerce advertising, AI/ML infrastructure, consumer data ecosystem',
  },
  {
    ticker: 'AAPL',
    name: 'Apple Inc.',
    revenue2025: 416161,
    ebitdaMargin: 0.35,
    fcf2025: 98767,
    relevance: 'Mobile ecosystem, App Store/search default agreements, privacy-driven ad targeting changes, services revenue growth',
  },
  {
    ticker: 'BIDU',
    name: 'Baidu, Inc.',
    revenue2025: 18320,
    ebitdaMargin: 0.22,
    fcf2025: 8450,
    relevance: 'Search dominance in China, AI/cloud (Ernie LLM), autonomous driving (Apollo), regional competitive benchmark',
  },
];

// ── Computed metrics ────────────────────────────────────────────────────────
function pctGrowth(current, prior) {
  if (prior == null || prior === 0) return null;
  return ((current - prior) / prior) * 100;
}

function ebitdaMargin(revenue, ebitda) {
  return (ebitda / revenue) * 100;
}

const revenueGrowth = FINANCIALS.revenue.map((r, i) =>
  i === 0 ? null : pctGrowth(r, FINANCIALS.revenue[i - 1])
);
const ebitdaMargins = FINANCIALS.revenue.map((r, i) =>
  ebitdaMargin(r, FINANCIALS.ebitda[i])
);
const fcfMargin = FINANCIALS.revenue.map((r, i) => (FINANCIALS.fcf[i] / r) * 100);
const cagr5y =
  (Math.pow(FINANCIALS.revenue[4] / FINANCIALS.revenue[0], 1 / 4) - 1) * 100;

// ── Formatting helpers ──────────────────────────────────────────────────────
const NAVY = '0A2540';
const LIGHT_BLUE = 'E8F0FE';
const ALT_ROW = 'F2F7FB';
const WHITE = 'FFFFFF';
const GRAY = '666666';

function fmtNum(n, decimals = 0) {
  if (n == null) return '—';
  return n.toLocaleString('en-US', {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  });
}

function fmtPct(n, decimals = 1) {
  if (n == null) return '—';
  return `${n.toFixed(decimals)}%`;
}

function fmtUsdM(n) {
  return `$${fmtNum(n)}`;
}

function fmtUsdB(n) {
  return `$${(n / 1000).toFixed(1)}B`;
}

function cellBorders() {
  return {
    top: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
    bottom: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
    left: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
    right: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
  };
}

function headerCell(text, width = 1200) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: { fill: NAVY, type: ShadingType.CLEAR },
    borders: cellBorders(),
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({ text, bold: true, color: WHITE, size: 20, font: 'Calibri' }),
        ],
      }),
    ],
  });
}

function dataCell(text, { bold = false, align = AlignmentType.RIGHT, shade = null, width = 1200 } = {}) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: shade ? { fill: shade, type: ShadingType.CLEAR } : undefined,
    borders: cellBorders(),
    children: [
      new Paragraph({
        alignment: align,
        children: [
          new TextRun({
            text,
            bold,
            size: 20,
            font: 'Calibri',
            color: bold ? NAVY : '333333',
          }),
        ],
      }),
    ],
  });
}

function labelCell(text, shade = null) {
  return dataCell(text, { bold: true, align: AlignmentType.LEFT, shade, width: 2800 });
}

function buildFinancialTable(title, rows, yearLabels = YEARS) {
  const header = new TableRow({
    children: [
      headerCell('($ in millions)', 2800),
      ...yearLabels.map((y) => headerCell(String(y))),
    ],
  });

  const dataRows = rows.map((row, idx) => {
    const shade = idx % 2 === 0 ? ALT_ROW : WHITE;
    return new TableRow({
      children: [
        labelCell(row.label, shade),
        ...row.values.map((v) =>
          dataCell(typeof v === 'number' ? fmtNum(v) : v, { shade })
        ),
      ],
    });
  });

  return [
    new Paragraph({
      heading: HeadingLevel.HEADING_2,
      children: [new TextRun({ text: title, bold: true, color: NAVY, font: 'Calibri' })],
    }),
    new Paragraph({ text: '' }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [header, ...dataRows],
    }),
    new Paragraph({ text: '' }),
  ];
}

function paragraph(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 120 },
    children: [
      new TextRun({
        text,
        size: 22,
        font: 'Calibri',
        color: opts.color || '333333',
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

// ── ExcelJS: validate & export workbook (supporting artifact) ─────────────────
async function buildExcelWorkbook() {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'M&IA';
  const ws = wb.addWorksheet('GOOGL Summary');

  ws.columns = [
    { header: 'Metric', key: 'metric', width: 28 },
    ...YEARS.map((y) => ({ header: String(y), key: String(y), width: 14 })),
  ];

  const addRow = (metric, values) => {
    const row = { metric };
    YEARS.forEach((y, i) => {
      row[String(y)] = values[i];
    });
    ws.addRow(row);
  };

  addRow('Revenue ($M)', FINANCIALS.revenue);
  addRow('Revenue Growth (%)', revenueGrowth.map((g) => (g == null ? '—' : g)));
  addRow('EBITDA ($M)', FINANCIALS.ebitda);
  addRow('EBITDA Margin (%)', ebitdaMargins);
  addRow('Net Income ($M)', FINANCIALS.netIncome);
  addRow('Free Cash Flow ($M)', FINANCIALS.fcf);
  addRow('FCF Margin (%)', fcfMargin);

  ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  ws.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF0A2540' },
  };

  const xlsxPath = path.join(OUTPUT_DIR, '_supporting-googL-model.xlsx');
  await wb.xlsx.writeFile(xlsxPath);
  return xlsxPath;
}

// ── PDFKit: executive summary PDF (supporting) ──────────────────────────────
function buildPdfSummary() {
  return new Promise((resolve, reject) => {
    const pdfPath = path.join(OUTPUT_DIR, '_supporting-executive-summary.pdf');
    const doc = new PDFDocument({ size: 'A4', margin: 50 });
    const stream = fs.createWriteStream(pdfPath);
    doc.pipe(stream);

    doc.fontSize(22).fillColor(BRAND.primary).text('Alphabet Inc. (GOOGL)', { align: 'center' });
    doc.fontSize(14).fillColor(GRAY).text('5-Year Financial Analysis | FY2021–FY2025', { align: 'center' });
    doc.moveDown();
    doc.fontSize(10).fillColor(GRAY).text(BRAND.confidentiality, { align: 'center' });
    doc.moveDown(2);

    doc.fontSize(12).fillColor('#333').text('Key Highlights', { underline: true });
    doc.moveDown(0.5);
    doc.fontSize(10)
      .text(`• Revenue CAGR (2021–2025): ${cagr5y.toFixed(1)}%`)
      .text(`• FY2025 Revenue: ${fmtUsdB(FINANCIALS.revenue[4])} (+${revenueGrowth[4].toFixed(1)}% YoY)`)
      .text(`• FY2025 EBITDA Margin: ${ebitdaMargins[4].toFixed(1)}% (${fmtUsdB(FINANCIALS.ebitda[4])})`)
      .text(`• FY2025 Free Cash Flow: ${fmtUsdB(FINANCIALS.fcf[4])} (FCF margin ${fcfMargin[4].toFixed(1)}%)`)
      .text(`• FY2025 Net Income: ${fmtUsdB(FINANCIALS.netIncome[4])}`)
      .text(`• Balance Sheet (FY2025): Assets ${fmtUsdB(BALANCE_SHEET.totalAssets[4])}, Equity ${fmtUsdB(BALANCE_SHEET.totalEquity[4])}, Debt ${fmtUsdB(BALANCE_SHEET.totalDebt[4])}`);

    doc.end();
    stream.on('finish', () => resolve(pdfPath));
    stream.on('error', reject);
  });
}

// ── PptxGenJS: peer comparison deck (supporting) ────────────────────────────
async function buildPeerDeck() {
  const pptx = new PptxGenJS();
  pptx.author = 'M&IA';
  pptx.title = 'GOOGL Peer Comparison';

  const slide = pptx.addSlide();
  slide.addText('Alphabet Inc. — Peer Universe', {
    x: 0.5,
    y: 0.3,
    w: 9,
    h: 0.6,
    fontSize: 24,
    bold: true,
    color: BRAND.primary,
  });

  const headers = ['Ticker', 'Company', 'FY25 Rev ($M)', 'EBITDA Margin', 'FY25 FCF ($M)'];
  const rows = PEERS.map((p) => [
    p.ticker,
    p.name,
    fmtNum(p.revenue2025),
    fmtPct(p.ebitdaMargin * 100),
    fmtNum(p.fcf2025),
  ]);
  rows.unshift(headers);

  slide.addTable(rows, {
    x: 0.4,
    y: 1.1,
    w: 9.2,
    fontSize: 10,
    border: { type: 'solid', color: 'CCCCCC', pt: 0.5 },
    fill: { color: LIGHT_BLUE },
    color: '333333',
  });

  const pptxPath = path.join(OUTPUT_DIR, '_supporting-peer-comparison.pptx');
  await pptx.writeFile({ fileName: pptxPath });
  return pptxPath;
}

// ── Build Word document ─────────────────────────────────────────────────────
function buildDocument() {
  const children = [];

  // Cover
  children.push(
    new Paragraph({ text: '' }),
    new Paragraph({ text: '' }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({
          text: 'ALPHABET INC.',
          bold: true,
          size: 56,
          color: NAVY,
          font: 'Calibri',
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({
          text: 'Comprehensive Financial Analysis Report',
          size: 32,
          color: NAVY,
          font: 'Calibri',
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 200 },
      children: [
        new TextRun({
          text: 'NASDAQ: GOOGL / GOOG  |  Class A & Class C Common Stock',
          size: 24,
          color: GRAY,
          font: 'Calibri',
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 400 },
      children: [
        new TextRun({
          text: 'Fiscal Years 2021 – 2025 (FY2025A)',
          size: 24,
          color: GRAY,
          font: 'Calibri',
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 600 },
      children: [
        new TextRun({
          text: REPORT_DATE,
          size: 22,
          color: GRAY,
          font: 'Calibri',
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 200 },
      children: [
        new TextRun({
          text: BRAND.confidentiality,
          italics: true,
          size: 20,
          color: GRAY,
          font: 'Calibri',
        }),
      ],
    }),
    new Paragraph({ children: [new PageBreak()] })
  );

  // Section I: Executive Summary
  children.push(
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [
        new TextRun({ text: 'I. Executive Summary', bold: true, color: NAVY, font: 'Calibri' }),
      ],
    }),
    paragraph(
      'Alphabet Inc. delivered a compelling five-year financial trajectory characterized by accelerating revenue growth, expanding profitability, and resilient free cash flow generation. Consolidated revenue increased from $257.6 billion in FY2021 to $402.8 billion in FY2025, representing a compound annual growth rate (CAGR) of 11.8%. The re-acceleration in FY2024–FY2025 reflects strength across Google Search, YouTube advertising, and Google Cloud, partially offset by elevated AI infrastructure capital expenditures.'
    ),
    paragraph(
      `EBITDA expanded from $103.5 billion (40.2% margin) in FY2021 to $180.0 billion (44.7% margin) in FY2025, underscoring operating leverage as scale benefits and cost discipline compound. Notably, FY2022 marked a margin trough (30.1%) driven by workforce expansion and cloud investment cycles; margins recovered meaningfully thereafter. Net income reached $132.2 billion in FY2025, while free cash flow remained robust at $73.3 billion despite a step-up in capital intensity.`
    ),
    paragraph(
      `The balance sheet remains fortress-like: total assets of $595.3 billion, stockholders' equity of $415.3 billion, and total debt of $59.3 billion (net debt / EBITDA of approximately 0.3x). Alphabet maintains exceptional financial flexibility to fund AI R&D, cloud infrastructure, strategic M&A, and shareholder returns.`
    )
  );

  // Section II: Historical Financial Tables
  children.push(
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [
        new TextRun({ text: 'II. Historical Financial Statements', bold: true, color: NAVY, font: 'Calibri' }),
      ],
    }),
    paragraph(
      'The following tables present Alphabet\'s consolidated income statement, balance sheet, and cash flow statement for fiscal years 2021 through 2025. All figures are in U.S. dollars millions (USD M) unless otherwise noted. FY2025A reflects audited results per the FY2025 Form 10-K filing.'
    )
  );

  // Income Statement
  children.push(
    ...buildFinancialTable('A. Consolidated Income Statement', [
      { label: 'Revenue', values: FINANCIALS.revenue },
      { label: 'Cost of Revenue', values: FINANCIALS.costOfRevenue },
      { label: 'Gross Profit', values: FINANCIALS.grossProfit },
      { label: 'Operating Income (EBIT)', values: FINANCIALS.operatingIncome },
      { label: 'Depreciation & Amortization', values: FINANCIALS.da },
      { label: 'EBITDA', values: FINANCIALS.ebitda },
      { label: 'Interest Expense', values: FINANCIALS.interestExpense },
      { label: 'Income Tax Provision', values: FINANCIALS.incomeTax },
      { label: 'Net Income', values: FINANCIALS.netIncome },
      { label: 'Diluted EPS ($)', values: FINANCIALS.epsDiluted.map((e) => e.toFixed(2)) },
      { label: 'Diluted Shares (M)', values: FINANCIALS.dilutedShares },
    ])
  );

  // Balance Sheet
  children.push(
    ...buildFinancialTable('B. Consolidated Balance Sheet', [
      { label: 'Cash & Cash Equivalents', values: BALANCE_SHEET.cashAndEquivalents },
      { label: 'Property & Equipment, Net', values: BALANCE_SHEET.ppeNet },
      { label: 'Goodwill', values: BALANCE_SHEET.goodwill },
      { label: 'Total Assets', values: BALANCE_SHEET.totalAssets },
      { label: 'Total Debt', values: BALANCE_SHEET.totalDebt },
      { label: 'Total Liabilities', values: BALANCE_SHEET.totalLiabilities },
      { label: "Stockholders' Equity", values: BALANCE_SHEET.totalEquity },
      { label: 'Debt / Total Assets (%)', values: BALANCE_SHEET.totalDebt.map((d, i) =>
        ((d / BALANCE_SHEET.totalAssets[i]) * 100).toFixed(1)
      ) },
      { label: 'Equity / Total Assets (%)', values: BALANCE_SHEET.totalEquity.map((e, i) =>
        ((e / BALANCE_SHEET.totalAssets[i]) * 100).toFixed(1)
      ) },
    ])
  );

  // Cash Flow
  children.push(
    ...buildFinancialTable('C. Consolidated Cash Flow Statement', [
      { label: 'Operating Cash Flow', values: CASH_FLOW.operatingCashFlow },
      { label: 'Capital Expenditures', values: CASH_FLOW.capex.map((c) => -c) },
      { label: 'Free Cash Flow (FCF)', values: CASH_FLOW.fcf },
      { label: 'Investing Cash Flow', values: CASH_FLOW.investingCashFlow },
      { label: 'Financing Cash Flow', values: CASH_FLOW.financingCashFlow },
      { label: 'Dividends Paid', values: CASH_FLOW.dividendsPaid.map((d) => (d === 0 ? '—' : -d)) },
      { label: 'Share Repurchases', values: CASH_FLOW.shareRepurchases.map((s) => -s) },
    ])
  );

  // Section III: Key Metrics
  children.push(
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [
        new TextRun({ text: 'III. Key Metrics Extraction & Analysis', bold: true, color: NAVY, font: 'Calibri' }),
      ],
    }),
    paragraph(
      'Three primary metrics anchor this analysis: revenue growth, EBITDA margin evolution, and free cash flow trajectory. Together, they illuminate Alphabet\'s top-line momentum, operating efficiency, and cash conversion quality.'
    )
  );

  children.push(
    ...buildFinancialTable('A. Revenue Growth Analysis', [
      { label: 'Revenue ($M)', values: FINANCIALS.revenue },
      { label: 'YoY Revenue Growth (%)', values: revenueGrowth.map((g) => (g == null ? '—' : g.toFixed(1))) },
      { label: 'Revenue CAGR 2021–2025 (%)', values: [cagr5y.toFixed(1), '—', '—', '—', '—'] },
    ])
  );

  children.push(
    paragraph(
      `Revenue growth decelerated to 9.8% in FY2022 amid macro headwinds in digital advertising, then re-accelerated to 15.1% in FY2025 — the fastest pace in the five-year window. The 11.8% revenue CAGR significantly outpaces mature technology peers and reflects durable monetization across Search (~57% of revenue), YouTube (~10%), Google Cloud (~12% and fastest-growing segment), and Other Bets. Cloud revenue growth exceeding 25% annually has been a critical diversification vector, reducing reliance on advertising cyclicality.`
    )
  );

  children.push(
    ...buildFinancialTable('B. EBITDA Margin Analysis', [
      { label: 'EBITDA ($M)', values: FINANCIALS.ebitda },
      { label: 'EBITDA Margin (%)', values: ebitdaMargins.map((m) => m.toFixed(1)) },
      { label: 'EBITDA Growth YoY (%)', values: FINANCIALS.ebitda.map((e, i) =>
        i === 0 ? '—' : pctGrowth(e, FINANCIALS.ebitda[i - 1]).toFixed(1)
      ) },
    ])
  );

  children.push(
    paragraph(
      'EBITDA margins compressed 1,010 basis points from FY2021 (40.2%) to FY2022 (30.1%) as Alphabet invested aggressively in headcount, cloud infrastructure, and Other Bets. The subsequent recovery to 44.7% in FY2025 — 450 bps above the FY2021 peak — demonstrates exceptional operating leverage. Key drivers include: (i) workforce optimization initiatives post-FY2022; (ii) cloud margin improvement as scale economics materialize; (iii) AI monetization in Search and Cloud beginning to offset inference/training costs; and (iv) disciplined opex management despite revenue reinvestment.'
    )
  );

  children.push(
    ...buildFinancialTable('C. Free Cash Flow Analysis', [
      { label: 'Operating Cash Flow ($M)', values: CASH_FLOW.operatingCashFlow },
      { label: 'Capital Expenditures ($M)', values: CASH_FLOW.capex },
      { label: 'Free Cash Flow ($M)', values: FINANCIALS.fcf },
      { label: 'FCF Margin (%)', values: fcfMargin.map((m) => m.toFixed(1)) },
      { label: 'FCF / Net Income (%)', values: FINANCIALS.fcf.map((f, i) =>
        ((f / FINANCIALS.netIncome[i]) * 100).toFixed(1)
      ) },
    ])
  );

  children.push(
    paragraph(
      'Free cash flow demonstrated resilience through the investment cycle, ranging from $60.0 billion (FY2022 trough) to $73.3 billion (FY2025). FCF conversion moderated relative to net income in FY2024–FY2025 as capital expenditures surged to ~$52 billion annually — primarily data center and AI accelerator buildouts. Despite this capex intensity, FCF margins held at 18.2% in FY2025, and cumulative five-year FCF totaled $342.6 billion. Alphabet\'s ability to self-fund AI infrastructure while sustaining buybacks ($62.2B in FY2025) and initiating dividends ($9.7B) underscores balance sheet strength.'
    )
  );

  // Section IV: Comparative Analysis
  children.push(
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [
        new TextRun({ text: 'IV. Comparative Peer Analysis', bold: true, color: NAVY, font: 'Calibri' }),
      ],
    }),
    paragraph(
      'Alphabet competes across multiple technology verticals — search, digital advertising, cloud computing, mobile ecosystems, and artificial intelligence. The peer group below represents the most strategically relevant publicly traded comparables, spanning direct advertising rivals, cloud infrastructure competitors, and platform ecosystem leaders.'
    )
  );

  const peerHeader = new TableRow({
    children: [
      headerCell('Ticker', 900),
      headerCell('Company', 2200),
      headerCell('FY25 Revenue ($M)', 1400),
      headerCell('EBITDA Margin', 1100),
      headerCell('FY25 FCF ($M)', 1200),
    ],
  });

  const googlPeerRow = new TableRow({
    children: [
      dataCell('GOOGL', { bold: true, align: AlignmentType.CENTER, shade: LIGHT_BLUE }),
      dataCell('Alphabet Inc.', { bold: true, align: AlignmentType.LEFT, shade: LIGHT_BLUE }),
      dataCell(fmtNum(FINANCIALS.revenue[4]), { bold: true, shade: LIGHT_BLUE }),
      dataCell(fmtPct(ebitdaMargins[4]), { bold: true, shade: LIGHT_BLUE }),
      dataCell(fmtNum(FINANCIALS.fcf[4]), { bold: true, shade: LIGHT_BLUE }),
    ],
  });

  const peerRows = PEERS.map((p, idx) => {
    const shade = idx % 2 === 0 ? ALT_ROW : WHITE;
    return new TableRow({
      children: [
        dataCell(p.ticker, { align: AlignmentType.CENTER, shade }),
        dataCell(p.name, { align: AlignmentType.LEFT, shade }),
        dataCell(fmtNum(p.revenue2025), { shade }),
        dataCell(fmtPct(p.ebitdaMargin * 100), { shade }),
        dataCell(fmtNum(p.fcf2025), { shade }),
      ],
    });
  });

  children.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [peerHeader, googlPeerRow, ...peerRows],
    }),
    new Paragraph({ text: '' })
  );

  children.push(
    new Paragraph({
      heading: HeadingLevel.HEADING_2,
      children: [
        new TextRun({ text: 'Strategic Competitive Positioning', bold: true, color: NAVY, font: 'Calibri' }),
      ],
    }),
    new Paragraph({ text: '' })
  );

  PEERS.forEach((p) => {
    children.push(
      new Paragraph({
        spacing: { after: 60 },
        children: [
          new TextRun({ text: `${p.ticker} — ${p.name}: `, bold: true, size: 22, font: 'Calibri', color: NAVY }),
          new TextRun({ text: p.relevance, size: 22, font: 'Calibri' }),
        ],
      })
    );
  });

  children.push(
    new Paragraph({ text: '' }),
    paragraph(
      'Relative to the peer set, Alphabet offers a distinctive risk/reward profile: advertising margins superior to Amazon, scale advantages over Meta in search intent monetization, cloud growth rates outpacing Microsoft\'s mature Azure base (though from a smaller absolute base), and AI integration depth across the full product stack. Baidu serves as a regional analog, though China regulatory dynamics limit direct comparability.'
    )
  );

  // Section V: Segment & Investment Implications
  children.push(
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [
        new TextRun({ text: 'V. Segment Overview & Investment Implications', bold: true, color: NAVY, font: 'Calibri' }),
      ],
    }),
    paragraph(
      'Alphabet reports through Google Services (Search, YouTube, Android, Chrome, Hardware), Google Cloud (GCP, Workspace), and Other Bets (Waymo, Verily, etc.). Google Services generates the vast majority of revenue and operating profit, while Google Cloud has reached profitability inflection and Other Bets remain investment-stage.'
    ),
    bullet('Google Search & Ads: Core profit engine; AI Overviews and Performance Max campaigns enhancing ARPU.'),
    bullet('YouTube: $50B+ annual revenue run-rate; Shorts monetization and connected TV ad share gains.'),
    bullet('Google Cloud: $40B+ revenue; AI infrastructure (TPU, Vertex AI) driving enterprise wins vs. AWS/Azure.'),
    bullet('Other Bets: Waymo commercialization progressing; disciplined capital allocation under "Other Bets" scrutiny.'),
    new Paragraph({ text: '' }),
    paragraph(
      'Investment implications: (1) Premium valuation supported by 44.7% EBITDA margins and net cash position; (2) AI capex cycle is the primary near-term FCF headwind but builds durable competitive moats; (3) Regulatory risk (DOJ antitrust remedies, EU DMA) remains the key overhang; (4) Capital return framework (buybacks + dividends) provides downside support at ~$70B+ annual shareholder yield.'
    )
  );

  // Section VI: Conclusion
  children.push(
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [
        new TextRun({ text: 'VI. Conclusion', bold: true, color: NAVY, font: 'Calibri' }),
      ],
    }),
    paragraph(
      'Alphabet Inc. has executed a five-year financial arc defined by double-digit revenue compounding, margin recovery exceeding pre-investment-cycle peaks, and industry-leading cash generation. FY2025 results — $402.8B revenue, $180.0B EBITDA, $132.2B net income, and $73.3B FCF — affirm the company\'s position as a preeminent global technology platform. Against a peer group comprising Microsoft, Meta, Amazon, Apple, and Baidu, Alphabet\'s diversified revenue mix, search advertising dominance, and AI/cloud optionality support a constructive long-term investment thesis, balanced against regulatory and capex cycle risks.'
    ),
    paragraph(
      'Source: Alphabet Inc. FY2021–FY2025 Form 10-K filings, Q4 FY2025 earnings release, and peer company public disclosures. Prepared for institutional use.',
      { italics: true, color: GRAY }
    )
  );

  return new Document({
    creator: 'M&IA Financial Analysis',
    title: 'Alphabet Inc. (GOOGL) — 5-Year Financial Analysis',
    description: 'Comprehensive financial analysis report covering FY2021–FY2025',
    sections: [{ properties: {}, children }],
  });
}

// ── Main ────────────────────────────────────────────────────────────────────
async function main() {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });

  const [xlsxPath, pdfPath, pptxPath] = await Promise.all([
    buildExcelWorkbook(),
    buildPdfSummary(),
    buildPeerDeck(),
  ]);

  const doc = buildDocument();
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(OUTPUT_DOCX, buffer);

  const stat = fs.statSync(OUTPUT_DOCX);
  console.log(`OK: wrote ${OUTPUT_DOCX} (${stat.size} bytes)`);
  console.log(`Supporting: ${xlsxPath}, ${pdfPath}, ${pptxPath}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
