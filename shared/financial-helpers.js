/**
 * Financial helper functions for M&IA document generation.
 * Used by generators to compute standard financial metrics.
 */

export function wacc({ equityWeight, costOfEquity, debtWeight, costOfDebt, taxRate }) {
  return equityWeight * costOfEquity + debtWeight * costOfDebt * (1 - taxRate);
}

export function npv(cashFlows, discountRate) {
  return cashFlows.reduce((acc, cf, i) => acc + cf / Math.pow(1 + discountRate, i + 1), 0);
}

export function irr(cashFlows, guess = 0.1) {
  const maxIter = 100;
  const tolerance = 1e-7;
  let rate = guess;
  for (let i = 0; i < maxIter; i++) {
    let npvVal = 0;
    let dNpv = 0;
    for (let t = 0; t < cashFlows.length; t++) {
      npvVal += cashFlows[t] / Math.pow(1 + rate, t);
      dNpv -= t * cashFlows[t] / Math.pow(1 + rate, t + 1);
    }
    if (Math.abs(npvVal) < tolerance) return rate;
    rate -= npvVal / dNpv;
  }
  return rate;
}

export function terminalValue({ fcf, growthRate, discountRate }) {
  return fcf * (1 + growthRate) / (discountRate - growthRate);
}

export function evToEbitda(ev, ebitda) {
  if (!ebitda || ebitda === 0) return null;
  return ev / ebitda;
}

export function debtToEbitda(netDebt, ebitda) {
  if (!ebitda || ebitda === 0) return null;
  return netDebt / ebitda;
}

export function formatEur(value, decimals = 1) {
  if (value >= 1e9) return `€${(value / 1e9).toFixed(decimals)}bn`;
  if (value >= 1e6) return `€${(value / 1e6).toFixed(decimals)}m`;
  if (value >= 1e3) return `€${(value / 1e3).toFixed(decimals)}k`;
  return `€${value.toFixed(0)}`;
}
