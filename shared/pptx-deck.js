// shared/pptx-deck.js
// Declarative PowerPoint deck composer (16:9, M&IA master).
//
// buildDeck(spec) → Buffer where spec = {
//   deckTitle, deckSubtitle?, project?, date?, confidential?, lang?,
//   slides: [
//     { type: 'section', title }
//     { type: 'content', title, bullets?, paragraphs?, table?, twoCol?: { left, right } }
//     { type: 'table', title, table: { headers, rows } }
//     { type: 'chart', title, chart: { kind, ... } }   kind: bar|line|pie|waterfall|footballField|heatmap
//     { type: 'facts', title, facts: [[label, value]] }
//     { type: 'kpi', title, kpis: [{ label, value, sub? }], bullets? }
//     { type: 'toc', title, items: [string] }
//     { type: 'matrix', title, matrix: { xLabel, yLabel, points: [{ name, x: 0-10, y: 0-10, highlight? }] } }
//     { type: 'custom', title, build(pptx, slide) }
//   ]
// }

import PptxGenJS from 'pptxgenjs';
import { BRAND } from './branding.js';
import {
  addBarChart, addLineChart, addPieChart, addWaterfallChart,
  addFootballField, addRiskHeatmap,
} from './chart-engine.js';

const SLIDE_W = 13.33;
const SLIDE_H = 7.5;
const MARGIN = 0.55;
const CONTENT_W = SLIDE_W - MARGIN * 2;

function confidentialityStamp(slide, lang) {
  slide.addText(lang === 'fr' ? BRAND.confidentialityFr : BRAND.confidentiality, {
    x: MARGIN, y: SLIDE_H - 0.42, w: 4, h: 0.3,
    fontSize: 8, italic: true, color: BRAND.medGray, fontFace: BRAND.font,
  });
}

function slideTitleBar(slide, text, lang, pageNum) {
  slide.addShape('rect', { x: 0, y: 0, w: SLIDE_W, h: 0.12, fill: { color: BRAND.accent } });
  slide.addText(text || '', {
    x: MARGIN, y: 0.3, w: CONTENT_W - 1, h: 0.6,
    fontSize: 22, bold: true, color: BRAND.navy, fontFace: BRAND.font,
  });
  slide.addText(String(pageNum), {
    x: SLIDE_W - 0.9, y: SLIDE_H - 0.42, w: 0.5, h: 0.3,
    fontSize: 9, color: BRAND.medGray, align: 'right', fontFace: BRAND.font,
  });
  confidentialityStamp(slide, lang);
}

function tableOptionsFromSpec(table) {
  const rows = [];
  rows.push(table.headers.map((h) => ({
    text: String(h),
    options: { bold: true, color: BRAND.white, fill: { color: BRAND.navy }, fontSize: 11, valign: 'middle' },
  })));
  for (let ri = 0; ri < table.rows.length; ri++) {
    rows.push(table.rows[ri].map((v, ci) => {
      const isNumeric = typeof v === 'number';
      const text = isNumeric ? v.toLocaleString('fr-FR', { maximumFractionDigits: 1 }) : String(v ?? '—');
      const negative = ci > 0 && (isNumeric ? v < 0 : /^\(|^-\d/.test(text.trim()));
      return {
        text,
        options: {
          fontSize: 10,
          color: negative ? BRAND.negative : BRAND.text,
          bold: ci === 0,
          align: ci === 0 ? 'left' : 'right',
          fill: ri % 2 === 1 ? { color: BRAND.lightGray } : { color: BRAND.white },
          valign: 'middle',
        },
      };
    }));
  }
  return rows;
}

function addBullets(slide, bullets, { x = MARGIN, y = 1.2, w = CONTENT_W, h = SLIDE_H - 2 } = {}) {
  const items = bullets.map((b) => {
    const text = typeof b === 'string' ? b : b.text;
    const level = typeof b === 'string' ? 0 : (b.level || 0);
    return {
      text,
      options: {
        bullet: { characterCode: '2022', indent: 18 },
        indentLevel: level,
        fontSize: level === 0 ? 14 : 12,
        color: BRAND.text,
        fontFace: BRAND.font,
        paraSpaceAfter: 8,
      },
    };
  });
  slide.addText(items, { x, y, w, h, valign: 'top' });
}

const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, Number(v) || 0));

/** KPI cards row(s): up to 4 cards per row. Returns height consumed. */
function renderKpiCards(slide, kpis, { x, y, w }) {
  const items = kpis.slice(0, 8);
  const perRow = Math.min(4, items.length);
  const gap = 0.3;
  const cardW = (w - gap * (perRow - 1)) / perRow;
  const cardH = 1.5;
  items.forEach((k, i) => {
    const row = Math.floor(i / perRow);
    const col = i % perRow;
    const cx = x + col * (cardW + gap);
    const cy = y + row * (cardH + gap);
    slide.addShape('roundRect', {
      x: cx, y: cy, w: cardW, h: cardH, rectRadius: 0.05,
      fill: { color: BRAND.lightBlue }, line: { color: BRAND.accent, width: 1 },
    });
    slide.addText(String(k.value ?? '—'), {
      x: cx, y: cy + 0.15, w: cardW, h: 0.62,
      align: 'center', fontSize: 24, bold: true, color: BRAND.navy, fontFace: BRAND.font,
    });
    slide.addText(String(k.label || ''), {
      x: cx, y: cy + 0.78, w: cardW, h: 0.35,
      align: 'center', fontSize: 11, color: BRAND.text, fontFace: BRAND.font,
    });
    if (k.sub) {
      slide.addText(String(k.sub), {
        x: cx, y: cy + 1.12, w: cardW, h: 0.3,
        align: 'center', fontSize: 9, italic: true, color: BRAND.medGray, fontFace: BRAND.font,
      });
    }
  });
  const rowsUsed = Math.ceil(items.length / perRow);
  return rowsUsed * cardH + (rowsUsed - 1) * gap;
}

/** 2x2 positioning matrix with named points on 0-10 axes. */
function renderMatrix(slide, matrix, area) {
  slide.addShape('rect', {
    x: area.x, y: area.y, w: area.w, h: area.h,
    fill: { color: BRAND.white }, line: { color: BRAND.medGray, width: 1 },
  });
  slide.addShape('line', {
    x: area.x + area.w / 2, y: area.y, w: 0, h: area.h,
    line: { color: 'CCCCCC', width: 0.75, dashType: 'dash' },
  });
  slide.addShape('line', {
    x: area.x, y: area.y + area.h / 2, w: area.w, h: 0,
    line: { color: 'CCCCCC', width: 0.75, dashType: 'dash' },
  });
  if (matrix.xLabel) {
    slide.addText(String(matrix.xLabel), {
      x: area.x, y: area.y + area.h + 0.08, w: area.w, h: 0.3,
      align: 'center', fontSize: 10, italic: true, color: BRAND.medGray, fontFace: BRAND.font,
    });
  }
  if (matrix.yLabel) {
    slide.addText(String(matrix.yLabel), {
      x: area.x - 1.55, y: area.y + area.h / 2 - 0.15, w: 3, h: 0.3,
      align: 'center', fontSize: 10, italic: true, color: BRAND.medGray, fontFace: BRAND.font, rotate: 270,
    });
  }
  const r = 0.14;
  for (const p of matrix.points || []) {
    const px = area.x + (clamp(p.x, 0, 10) / 10) * (area.w - 0.4) + 0.2;
    const py = area.y + area.h - ((clamp(p.y, 0, 10) / 10) * (area.h - 0.4) + 0.2);
    slide.addShape('ellipse', {
      x: px - r, y: py - r, w: r * 2, h: r * 2,
      fill: { color: p.highlight ? BRAND.accent : BRAND.medGray },
      line: { color: BRAND.white, width: 1 },
    });
    slide.addText(String(p.name || ''), {
      x: px - 1, y: py + r + 0.02, w: 2, h: 0.28,
      align: 'center', fontSize: 9, bold: Boolean(p.highlight),
      color: p.highlight ? BRAND.navy : BRAND.text, fontFace: BRAND.font,
    });
  }
}

function renderChart(pptx, slide, chart, area) {
  const { kind, ...rest } = chart;
  const opts = { x: area.x, y: area.y, w: area.w, h: area.h, ...rest };

  // Normalize: allow { categories, series: [{ name, values }] } shorthand —
  // inject categories as labels on each series for bar/line charts.
  if (Array.isArray(opts.series)) {
    opts.series = opts.series.map((s) => ({
      ...s,
      labels: s.labels || opts.categories || s.values?.map((_, i) => String(i + 1)) || [],
      values: s.values || [],
    }));
  }
  // Normalize football field: accept `items`/`label` aliases for `ranges`/`method`.
  if (kind === 'footballField') {
    const src = opts.ranges || opts.items || [];
    opts.ranges = src.map((r) => ({ method: r.method || r.label || '', low: r.low ?? 0, high: r.high ?? 0 }));
  }
  if (kind === 'waterfall') opts.steps = opts.steps || opts.items || [];
  if (kind === 'heatmap') opts.risks = opts.risks || opts.items || [];

  if (kind === 'bar') addBarChart(pptx, slide, opts);
  else if (kind === 'line') addLineChart(pptx, slide, opts);
  else if (kind === 'pie') addPieChart(pptx, slide, opts);
  else if (kind === 'waterfall') addWaterfallChart(pptx, slide, opts);
  else if (kind === 'footballField') addFootballField(pptx, slide, opts);
  else if (kind === 'heatmap') addRiskHeatmap(pptx, slide, opts);
}

export async function buildDeck(spec) {
  const lang = spec.lang || 'fr';
  const deckTitle = spec.deckTitle || spec.title || 'Presentation';
  const deckSubtitle = spec.deckSubtitle || spec.subtitle;
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'WIDE';
  pptx.author = BRAND.brandName;
  pptx.title = deckTitle;

  // --- Title slide ---
  const titleSlide = pptx.addSlide();
  titleSlide.background = { color: BRAND.navy };
  titleSlide.addShape('rect', { x: 0, y: SLIDE_H - 0.25, w: SLIDE_W, h: 0.25, fill: { color: BRAND.accent } });
  titleSlide.addText(BRAND.brandName, {
    x: MARGIN, y: 0.5, w: 4, h: 0.5, fontSize: 18, bold: true, color: BRAND.accent, fontFace: BRAND.font,
  });
  titleSlide.addText(deckTitle, {
    x: MARGIN, y: 2.6, w: CONTENT_W, h: 1.4, fontSize: 40, bold: true, color: BRAND.white, fontFace: BRAND.font,
  });
  if (deckSubtitle) {
    titleSlide.addText(deckSubtitle, {
      x: MARGIN, y: 4.0, w: CONTENT_W, h: 0.7, fontSize: 20, color: BRAND.lightBlue, fontFace: BRAND.font,
    });
  }
  if (spec.project) {
    titleSlide.addText(spec.project, {
      x: MARGIN, y: 4.7, w: CONTENT_W, h: 0.5, fontSize: 16, italic: true, color: BRAND.lightBlue, fontFace: BRAND.font,
    });
  }
  titleSlide.addText(spec.date || new Date().toLocaleDateString(lang === 'fr' ? 'fr-FR' : 'en-GB', { year: 'numeric', month: 'long' }), {
    x: MARGIN, y: 6.4, w: 5, h: 0.4, fontSize: 13, color: BRAND.medGray, fontFace: BRAND.font,
  });
  if (spec.confidential !== false) {
    titleSlide.addText((lang === 'fr' ? BRAND.confidentialityFr : BRAND.confidentiality).toUpperCase(), {
      x: SLIDE_W - 4.5, y: 6.4, w: 4, h: 0.4, fontSize: 11, bold: true, color: BRAND.negative, align: 'right', fontFace: BRAND.font,
    });
  }

  // --- Body slides ---
  let pageNum = 1;
  for (const s of spec.slides || []) {
    pageNum += 1;
    const slide = pptx.addSlide();

    if (s.type === 'section') {
      slide.background = { color: BRAND.darkBlue };
      slide.addText(s.title || '', {
        x: MARGIN, y: SLIDE_H / 2 - 0.6, w: CONTENT_W, h: 1.2,
        fontSize: 32, bold: true, color: BRAND.white, fontFace: BRAND.font,
      });
      slide.addShape('rect', { x: MARGIN, y: SLIDE_H / 2 + 0.7, w: 2.5, h: 0.08, fill: { color: BRAND.accent } });
      continue;
    }

    slideTitleBar(slide, s.title, lang, pageNum);
    let cursorY = 1.15;

    if (s.paragraphs?.length) {
      slide.addText(s.paragraphs.join('\n\n'), {
        x: MARGIN, y: cursorY, w: CONTENT_W, h: Math.min(2.2, 0.5 * s.paragraphs.length + 0.3),
        fontSize: 13, color: BRAND.text, fontFace: BRAND.font, valign: 'top',
      });
      cursorY += Math.min(2.2, 0.5 * s.paragraphs.length + 0.3) + 0.1;
    }

    if (s.type === 'toc' && s.items?.length) {
      s.items.slice(0, 10).forEach((item, i) => {
        const y = cursorY + 0.15 + i * 0.52;
        slide.addText(String(i + 1).padStart(2, '0'), {
          x: MARGIN, y, w: 0.7, h: 0.45,
          fontSize: 16, bold: true, color: BRAND.accent, fontFace: BRAND.font,
        });
        slide.addText(String(item), {
          x: MARGIN + 0.85, y: y + 0.03, w: CONTENT_W - 0.85, h: 0.45,
          fontSize: 14, color: BRAND.navy, fontFace: BRAND.font,
        });
      });
      continue;
    }

    if (s.type === 'kpi' && s.kpis?.length) {
      const usedH = renderKpiCards(slide, s.kpis, { x: MARGIN, y: cursorY, w: CONTENT_W });
      cursorY += usedH + 0.25;
      if (s.bullets?.length) addBullets(slide, s.bullets, { y: cursorY, h: SLIDE_H - cursorY - 0.7 });
      continue;
    }

    if (s.type === 'matrix' && s.matrix?.points?.length) {
      renderMatrix(slide, s.matrix, {
        x: MARGIN + 0.7, y: cursorY + 0.15,
        w: CONTENT_W - 1.4, h: SLIDE_H - cursorY - 1.5,
      });
      continue;
    }

    if (s.type === 'facts' && s.facts?.length) {
      const rows = s.facts.map((fact) => {
        const [label, value] = Array.isArray(fact) ? fact : [fact.label, fact.value];
        return [
          { text: String(label), options: { bold: true, color: BRAND.navy, fontSize: 12, fill: { color: BRAND.lightBlue } } },
          { text: String(value ?? '—'), options: { fontSize: 12, color: BRAND.text } },
        ];
      });
      slide.addTable(rows, { x: MARGIN, y: cursorY, w: CONTENT_W, colW: [CONTENT_W * 0.35, CONTENT_W * 0.65], border: { type: 'solid', color: 'CCCCCC', pt: 0.5 }, rowH: 0.4 });
    }

    if (s.twoCol) {
      const colW = (CONTENT_W - 0.4) / 2;
      if (s.twoCol.left?.bullets) addBullets(slide, s.twoCol.left.bullets, { x: MARGIN, y: cursorY, w: colW, h: SLIDE_H - cursorY - 0.7 });
      if (s.twoCol.left?.chart) renderChart(pptx, slide, s.twoCol.left.chart, { x: MARGIN, y: cursorY, w: colW, h: SLIDE_H - cursorY - 0.9 });
      if (s.twoCol.right?.bullets) addBullets(slide, s.twoCol.right.bullets, { x: MARGIN + colW + 0.4, y: cursorY, w: colW, h: SLIDE_H - cursorY - 0.7 });
      if (s.twoCol.right?.chart) renderChart(pptx, slide, s.twoCol.right.chart, { x: MARGIN + colW + 0.4, y: cursorY, w: colW, h: SLIDE_H - cursorY - 0.9 });
    } else {
      if (s.bullets?.length) {
        addBullets(slide, s.bullets, { y: cursorY, h: SLIDE_H - cursorY - 0.7 });
        cursorY += Math.min(3.5, s.bullets.length * 0.4 + 0.2);
      }
      // Accept table spec either nested ({ table: { headers, rows } }) or flat ({ headers, rows }).
      const tableSpec = s.table || (s.headers && s.rows ? { headers: s.headers, rows: s.rows } : null);
      if (tableSpec) {
        slide.addTable(tableOptionsFromSpec(tableSpec), {
          x: MARGIN, y: cursorY, w: CONTENT_W,
          border: { type: 'solid', color: 'CCCCCC', pt: 0.5 },
          autoPage: true, autoPageRepeatHeader: true,
        });
      }
      if (s.chart) {
        renderChart(pptx, slide, s.chart, { x: MARGIN, y: cursorY, w: CONTENT_W, h: SLIDE_H - cursorY - 0.9 });
      }
    }

    if (s.type === 'custom' && typeof s.build === 'function') {
      s.build(pptx, slide, { MARGIN, CONTENT_W, SLIDE_W, SLIDE_H });
    }
  }

  // --- Disclaimer slide ---
  if (spec.disclaimer) {
    const last = pptx.addSlide();
    slideTitleBar(last, lang === 'fr' ? 'Avertissement' : 'Disclaimer', lang, pageNum + 1);
    last.addText(spec.disclaimer, {
      x: MARGIN, y: 1.3, w: CONTENT_W, h: 4.5,
      fontSize: 11, italic: true, color: BRAND.medGray, fontFace: BRAND.font, valign: 'top',
    });
  }

  const out = await pptx.write({ outputType: 'nodebuffer' });
  return Buffer.from(out);
}
