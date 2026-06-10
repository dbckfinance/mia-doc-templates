// shared/financial-helpers.js
// Number formatting + core finance math used across all verticals.

const FR = 'fr-FR';

export function formatNumber(value, { decimals = 1, locale = FR } = {}) {
  if (value == null || Number.isNaN(Number(value))) return '—';
  return Number(value).toLocaleString(locale, {
    minimumFractionDigits: 0,
    maximumFractionDigits: decimals,
  });
}

export function formatCurrency(value, { currency = 'EUR', decimals = 1, compact = true } = {}) {
  if (value == null || Number.isNaN(Number(value))) return '—';
  const n = Number(value);
  const symbol = currency === 'EUR' ? '€' : currency === 'USD' ? '$' : currency === 'GBP' ? '£' : `${currency} `;
  if (compact) {
    const abs = Math.abs(n);
    if (abs >= 1e9) return `${symbol}${formatNumber(n / 1e9, { decimals })}Md`;
    if (abs >= 1e6) return `${symbol}${formatNumber(n / 1e6, { decimals })}m`;
    if (abs >= 1e3) return `${symbol}${formatNumber(n / 1e3, { decimals })}k`;
  }
  return `${symbol}${formatNumber(n, { decimals })}`;
}

export function formatPercent(value, { decimals = 1, alreadyPercent = false } = {}) {
  if (value == null || Number.isNaN(Number(value))) return '—';
  const pct = alreadyPercent ? Number(value) : Number(value) * 100;
  return `${formatNumber(pct, { decimals })}%`;
}

export function formatMultiple(value, { decimals = 1 } = {}) {
  if (value == null || Number.isNaN(Number(value))) return '—';
  return `${formatNumber(value, { decimals })}x`;
}

// ---------------------------------------------------------------------------
// Growth / margins
// ---------------------------------------------------------------------------

export function cagr(first, last, years) {
  if (!first || !last || !years || first <= 0 || last <= 0) return null;
  return Math.pow(last / first, 1 / years) - 1;
}

export function margin(numerator, denominator) {
  if (!denominator) return null;
  return numerator / denominator;
}

export function variance(current, previous) {
  if (previous == null || current == null || previous === 0) return null;
  return (current - previous) / Math.abs(previous);
}

// ---------------------------------------------------------------------------
// Valuation math
// ---------------------------------------------------------------------------

export function npv(rate, cashflows) {
  return cashflows.reduce((acc, cf, i) => acc + cf / Math.pow(1 + rate, i + 1), 0);
}

export function irr(cashflows, { guess = 0.1, maxIter = 200, tol = 1e-7 } = {}) {
  // Newton-Raphson with bisection fallback
  let rate = guess;
  for (let i = 0; i < maxIter; i++) {
    let f = 0;
    let df = 0;
    for (let t = 0; t < cashflows.length; t++) {
      f += cashflows[t] / Math.pow(1 + rate, t);
      df += (-t * cashflows[t]) / Math.pow(1 + rate, t + 1);
    }
    if (Math.abs(f) < tol) return rate;
    if (df === 0) break;
    rate -= f / df;
    if (rate <= -0.9999) rate = -0.9;
  }
  // Bisection fallback over [-0.9, 10]
  let lo = -0.9;
  let hi = 10;
  const npvAt = (r) => cashflows.reduce((a, cf, t) => a + cf / Math.pow(1 + r, t), 0);
  if (npvAt(lo) * npvAt(hi) > 0) return null;
  for (let i = 0; i < 200; i++) {
    const mid = (lo + hi) / 2;
    const v = npvAt(mid);
    if (Math.abs(v) < tol) return mid;
    if (npvAt(lo) * v < 0) hi = mid;
    else lo = mid;
  }
  return (lo + hi) / 2;
}

export function wacc({ equityValue, debtValue, costOfEquity, costOfDebt, taxRate = 0.25 }) {
  const total = (equityValue || 0) + (debtValue || 0);
  if (!total) return null;
  return (
    (equityValue / total) * costOfEquity +
    (debtValue / total) * costOfDebt * (1 - taxRate)
  );
}

export function terminalValue({ finalCashflow, growthRate, discountRate }) {
  if (discountRate <= growthRate) return null;
  return (finalCashflow * (1 + growthRate)) / (discountRate - growthRate);
}

export function moic(totalDistributions, totalInvested) {
  if (!totalInvested) return null;
  return totalDistributions / totalInvested;
}

// ---------------------------------------------------------------------------
// Table helpers
// ---------------------------------------------------------------------------

/**
 * Build a variance table spec from yearly financials:
 *   rows: [{ label, values: [..] }] → adds Δ% column vs prior year.
 */
export function buildVarianceTable(years, rows) {
  const headers = ['', ...years.map(String), 'Δ% (N vs N-1)'];
  const body = rows.map((row) => {
    const vals = row.values || [];
    const last = vals[vals.length - 1];
    const prev = vals[vals.length - 2];
    const delta = variance(last, prev);
    return [
      row.label,
      ...vals.map((v) => (row.isPercent ? formatPercent(v) : formatNumber(v))),
      delta == null ? '—' : formatPercent(delta),
    ];
  });
  return { headers, rows: body };
}

/** Returns rows of standard credit metrics from a financials object. */
export function creditMetrics(fin = {}) {
  const out = [];
  if (fin.netDebt != null && fin.ebitda) {
    out.push(['Net Debt / EBITDA', formatMultiple(fin.netDebt / fin.ebitda)]);
  }
  if (fin.ebitda != null && fin.interestExpense) {
    out.push(['EBITDA / Interest', formatMultiple(fin.ebitda / fin.interestExpense)]);
  }
  if (fin.totalDebt != null && fin.totalAssets) {
    out.push(['Debt / Total Assets', formatPercent(fin.totalDebt / fin.totalAssets)]);
  }
  if (fin.ffo != null && fin.totalDebt) {
    out.push(['FFO / Debt', formatPercent(fin.ffo / fin.totalDebt)]);
  }
  return out;
}
