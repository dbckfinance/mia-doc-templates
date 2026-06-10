// shared/chart-engine.js
// pptxgenjs chart builders: bar, line, waterfall (stacked-bar trick),
// scatter, pie, and the valuation football field.

import { BRAND } from './branding.js';

const PALETTE = [BRAND.navy, BRAND.accent, BRAND.medGray, BRAND.positive, BRAND.warning, BRAND.darkBlue];

const AXIS_OPTS = {
  catAxisLabelColor: BRAND.text,
  catAxisLabelFontSize: 10,
  valAxisLabelColor: BRAND.text,
  valAxisLabelFontSize: 10,
  valGridLine: { color: 'E0E0E0', style: 'solid', size: 1 },
  catGridLine: { style: 'none' },
};

/** Vertical bar chart. series = [{ name, labels, values }] */
export function addBarChart(pptx, slide, { x, y, w, h, series, title: chartTitle, barDir = 'col' }) {
  slide.addChart(pptx.ChartType.bar, series.map((s, i) => ({
    name: s.name,
    labels: s.labels,
    values: s.values,
  })), {
    x, y, w, h,
    barDir,
    chartColors: PALETTE,
    showTitle: Boolean(chartTitle),
    title: chartTitle,
    titleFontSize: 12,
    titleColor: BRAND.navy,
    showLegend: series.length > 1,
    legendPos: 'b',
    legendFontSize: 9,
    dataLabelFontSize: 9,
    showValue: series.length === 1,
    dataLabelColor: BRAND.text,
    ...AXIS_OPTS,
  });
}

/** Line chart for time series. */
export function addLineChart(pptx, slide, { x, y, w, h, series, title: chartTitle }) {
  slide.addChart(pptx.ChartType.line, series, {
    x, y, w, h,
    chartColors: PALETTE,
    lineSize: 2.5,
    lineSmooth: false,
    showTitle: Boolean(chartTitle),
    title: chartTitle,
    titleFontSize: 12,
    titleColor: BRAND.navy,
    showLegend: series.length > 1,
    legendPos: 'b',
    legendFontSize: 9,
    ...AXIS_OPTS,
  });
}

/** Pie / donut. data = { labels, values } */
export function addPieChart(pptx, slide, { x, y, w, h, labels, values, title: chartTitle, donut = true }) {
  slide.addChart(donut ? pptx.ChartType.doughnut : pptx.ChartType.pie, [{ name: chartTitle || 'Mix', labels, values }], {
    x, y, w, h,
    chartColors: PALETTE,
    showTitle: Boolean(chartTitle),
    title: chartTitle,
    titleFontSize: 12,
    titleColor: BRAND.navy,
    showLegend: true,
    legendPos: 'r',
    legendFontSize: 9,
    showPercent: true,
    dataLabelFontSize: 9,
  });
}

/**
 * Waterfall / bridge chart via stacked bars: invisible base + visible delta.
 * steps = [{ label, value }] where value may be +/-; first/last treated as totals.
 */
export function addWaterfallChart(pptx, slide, { x, y, w, h, steps, title: chartTitle }) {
  const labels = steps.map((s) => s.label);
  const base = [];
  const totals = [];
  const ups = [];
  const downs = [];
  let running = 0;

  steps.forEach((s, i) => {
    const isTotal = i === 0 || i === steps.length - 1 || s.isTotal;
    if (isTotal) {
      const val = i === 0 ? s.value : (s.value != null ? s.value : running);
      base.push(0); totals.push(val); ups.push(0); downs.push(0);
      running = val;
    } else if (s.value >= 0) {
      base.push(running); totals.push(0); ups.push(s.value); downs.push(0);
      running += s.value;
    } else {
      running += s.value;
      base.push(running); totals.push(0); ups.push(0); downs.push(-s.value);
    }
  });

  slide.addChart(pptx.ChartType.bar, [
    { name: '_base', labels, values: base },
    { name: 'Total', labels, values: totals },
    { name: 'Increase', labels, values: ups },
    { name: 'Decrease', labels, values: downs },
  ], {
    x, y, w, h,
    barDir: 'col',
    barGrouping: 'stacked',
    chartColors: ['FFFFFF', BRAND.navy, BRAND.positive, BRAND.negative],
    chartColorsOpacity: 100,
    showTitle: Boolean(chartTitle),
    title: chartTitle,
    titleFontSize: 12,
    titleColor: BRAND.navy,
    showLegend: false,
    showValue: false,
    ...AXIS_OPTS,
  });
}

/**
 * Valuation football field as horizontal floating bars.
 * ranges = [{ method, low, high }]; optional marker = current/offer price.
 */
export function addFootballField(pptx, slide, { x, y, w, h, ranges, title: chartTitle, currency = '€' }) {
  const labels = ranges.map((r) => r.method);
  const lows = ranges.map((r) => r.low);
  const spans = ranges.map((r) => Math.max(0, r.high - r.low));

  slide.addChart(pptx.ChartType.bar, [
    { name: '_low', labels, values: lows },
    { name: 'Range', labels, values: spans },
  ], {
    x, y, w, h,
    barDir: 'bar',
    barGrouping: 'stacked',
    chartColors: ['FFFFFF', BRAND.accent],
    showTitle: Boolean(chartTitle),
    title: chartTitle,
    titleFontSize: 12,
    titleColor: BRAND.navy,
    showLegend: false,
    showValue: false,
    valAxisLabelFormatCode: `${currency === '€' ? '#,##0\\ "€"' : '"' + currency + '"#,##0'}`,
    ...AXIS_OPTS,
  });

  // Annotate min-max as text under the chart
  const lo = Math.min(...ranges.map((r) => r.low));
  const hi = Math.max(...ranges.map((r) => r.high));
  slide.addText(`Fourchette globale: ${currency}${lo.toLocaleString('fr-FR')} — ${currency}${hi.toLocaleString('fr-FR')}`, {
    x, y: y + h + 0.05, w, h: 0.3, fontSize: 10, italic: true, color: BRAND.medGray, fontFace: BRAND.font,
  });
}

/** Risk heatmap (impact x likelihood grid) drawn with shapes. */
export function addRiskHeatmap(pptx, slide, { x, y, w, h, risks }) {
  const cols = 5;
  const rows = 5;
  const cw = w / cols;
  const ch = h / rows;
  const colorFor = (impact, likelihood) => {
    const score = impact * likelihood;
    if (score >= 15) return BRAND.negative;
    if (score >= 8) return BRAND.warning;
    return BRAND.positive;
  };
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const impact = rows - r;
      const likelihood = c + 1;
      slide.addShape(pptx.ShapeType.rect, {
        x: x + c * cw, y: y + r * ch, w: cw - 0.02, h: ch - 0.02,
        fill: { color: colorFor(impact, likelihood), transparency: 70 },
        line: { color: 'FFFFFF', width: 1 },
      });
    }
  }
  (risks || []).forEach((risk, i) => {
    const impact = Math.min(5, Math.max(1, risk.impact || 3));
    const likelihood = Math.min(5, Math.max(1, risk.likelihood || 3));
    slide.addText(String(risk.id || i + 1), {
      x: x + (likelihood - 1) * cw + cw / 2 - 0.15,
      y: y + (rows - impact) * ch + ch / 2 - 0.15,
      w: 0.3, h: 0.3,
      fontSize: 10, bold: true, color: 'FFFFFF', fontFace: BRAND.font,
      fill: { color: BRAND.navy }, shape: pptx.ShapeType.ellipse, align: 'center',
    });
  });
  slide.addText('Probabilité →', { x, y: y + h + 0.05, w: w / 2, h: 0.25, fontSize: 9, color: BRAND.medGray, fontFace: BRAND.font });
  slide.addText('↑ Impact', { x: x - 0.05, y: y - 0.3, w: 1.2, h: 0.25, fontSize: 9, color: BRAND.medGray, fontFace: BRAND.font });
}
