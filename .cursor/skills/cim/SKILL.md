---
name: M&IA CIM Methodology
when_to_use: Producing a Confidential Information Memorandum for a sell-side mandate.
---

# CIM Generation Methodology

## How to generate

Always use the deterministic generator: `node generators/ma/cim.js input.json`.
It produces a branded 16:9 PPTX with TOC, KPI cards, charts, competitive matrix
and disclaimer automatically. Fill the input.json as richly as possible — every
optional field you provide becomes a slide.

## Structure (produced by the generator)

1. **Sommaire** — auto TOC slide
2. **Synthèse exécutive** — KPI cards (CA, EBITDA, marge, TCAC) + opportunity overview
3. **Points clés d'investissement** — numbered highlights (5-8)
4. **Présentation de l'entreprise** — facts card, activité, historique, offre produits
5. **Marché** — overview + drivers, market size bar chart, competitive table, positioning matrix
6. **Modèle économique** — revenue split pie + geographic split pie, client concentration
7. **Équipe de management** — table with bios
8. **Performance financière** — CA/EBITDA bar chart, EBITDA margin line chart, full table
9. **Business plan** — projections bar chart + assumptions
10. **Stratégie / Risques** — strategy bullets, risks & mitigants table
11. **Transaction** — deal structure facts
12. **Annexes** — custom sections (each can carry its own chart/table)

## Input fields that unlock charts (FILL THEM)

- `financials.revenue` + `financials.ebitda` + `years` → CA/EBITDA bars + margin line + table
- `market.size: { labels, values, unit }` → market size bar chart
- `market.competitors[].x/y` (0-10) → competitive positioning matrix
- `revenueSplit` / `geographicSplit: { labels, values }` → pie charts (side by side)
- `businessPlan: { years, revenue, ebitda, assumptions }` → projections chart
- `kpis: [{ label, value, sub }]` → KPI cards (auto-computed from financials if absent)
- `clients.concentration: [{ name, share }]` → concentration table
- `risks: [{ risk, mitigant }]` → risk table

## Quality Standards — bulge-bracket level

- MANDATORY: at least 4 charts in any CIM (financials, market, revenue mix, projections)
- Never leave a slide with a title only — every slide has content
- All financial data must be sourced and dated; use €M for European deals
- If a data point is missing, estimate it and mark it `[à compléter]` — never invent silently
- Minimum 20 slides for a standard CIM; 30-60 for a full process

## Data Requirements

Gather or receive before generating:
- Company name, sector, geography, ownership
- 3 years of P&L (Revenue, EBITDA, Net Income minimum)
- Revenue split by activity and geography
- Market size data and named competitors
- Management team with bios
- Business plan projections (3-5 years)
