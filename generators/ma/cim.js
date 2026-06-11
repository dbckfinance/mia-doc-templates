// generators/ma/cim.js
// Confidential Information Memorandum (pptx, 30-60 slides)
// IB-grade: TOC, KPI cards, numbered investment highlights, market charts,
// competitive matrix, revenue mix pies, multi-chart financials, BP projections.

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildDeck } from '../../shared/pptx-deck.js';
import { disclaimerFor } from '../../shared/disclaimer-library.js';
import { formatPercent } from '../../shared/financial-helpers.js';

export const metadata = {
  id: 'cim',
  name: 'Confidential Information Memorandum',
  vertical: 'ma',
  outputType: 'pptx',
  estimatedPages: '30-60 slides',
  requiredInput: ['company'],
  optionalInput: [
    'sector', 'founded', 'headquarters', 'headcount', 'ownership',
    'businessDescription', 'businessLines', 'history', 'products',
    'investmentHighlights', 'kpis',
    'market',          // { overview, drivers[], size: { labels[], values[], unit? }, competitors[{ name, positioning, share, x?, y? }] }
    'revenueSplit',    // { labels[], values[], title? }
    'geographicSplit', // { labels[], values[] }
    'clients',         // { concentration: [{ name, share }], commentary? }
    'management',      // [{ name, role, bio }]
    'financials',      // { revenue[], ebitda[], netIncome[], capex[], netDebt[], workingCapital[] }
    'years',
    'businessPlan',    // { years[], revenue[], ebitda[], assumptions[], commentary }
    'strategy', 'risks', 'transaction', 'sections',
  ],
};

const fmtM = (v) => (typeof v === 'number' ? `${v.toLocaleString('fr-FR', { maximumFractionDigits: 1 })} m€` : v ?? '—');

function cagr(values) {
  if (!Array.isArray(values) || values.length < 2 || !values[0]) return null;
  const n = values.length - 1;
  return (values[values.length - 1] / values[0]) ** (1 / n) - 1;
}

export async function generate(input = {}) {
  const company = input.company || 'Société';
  const fin = input.financials || {};
  const years = (input.years || []).map(String);
  const bp = input.businessPlan || {};
  const slides = [];

  // --- Sommaire ---
  slides.push({
    type: 'toc',
    title: 'Sommaire',
    items: [
      'Synthèse exécutive',
      'Points clés d\'investissement',
      'Présentation de l\'entreprise',
      'Marché et positionnement concurrentiel',
      'Modèle économique et clients',
      'Équipe de management',
      'Performance financière historique',
      'Business plan et perspectives',
      'Transaction envisagée',
      'Annexes',
    ],
  });

  // --- 1. Executive summary ---
  slides.push({ type: 'section', title: 'Synthèse exécutive' });

  const lastIdx = (fin.revenue?.length || 0) - 1;
  const revCagr = cagr(fin.revenue);
  const autoKpis = [];
  if (lastIdx >= 0) {
    autoKpis.push({ label: `Chiffre d'affaires ${years[lastIdx] || ''}`, value: fmtM(fin.revenue[lastIdx]) });
    if (fin.ebitda?.[lastIdx] != null) {
      autoKpis.push({ label: `EBITDA ${years[lastIdx] || ''}`, value: fmtM(fin.ebitda[lastIdx]) });
      autoKpis.push({
        label: 'Marge EBITDA',
        value: formatPercent(fin.ebitda[lastIdx] / fin.revenue[lastIdx]),
      });
    }
    if (revCagr != null) autoKpis.push({ label: `TCAC CA ${years[0]}-${years[lastIdx]}`, value: formatPercent(revCagr) });
  }
  const kpis = input.kpis?.length ? input.kpis : autoKpis;
  if (kpis.length) {
    slides.push({
      type: 'kpi',
      title: 'Chiffres clés',
      kpis,
      bullets: input.executiveSummary ? [input.executiveSummary] : [],
    });
  }
  slides.push({
    type: 'content',
    title: "L'opportunité en bref",
    paragraphs: input.businessDescription ? [input.businessDescription] : [],
    bullets: input.investmentHighlights || [
      'Leader sur son marché avec des barrières à l\'entrée fortes',
      'Croissance soutenue et rentabilité supérieure aux pairs',
      'Potentiel de consolidation et de développement international',
    ],
  });

  // --- 2. Investment highlights (numbered) ---
  if (input.investmentHighlights?.length) {
    slides.push({ type: 'section', title: "Points clés d'investissement" });
    slides.push({
      type: 'content',
      title: "Pourquoi investir dans " + company,
      bullets: input.investmentHighlights.map((h, i) => `${i + 1}. ${h}`),
    });
  }

  // --- 3. Company overview ---
  slides.push({ type: 'section', title: "Présentation de l'entreprise" });
  slides.push({
    type: 'facts',
    title: "Carte d'identité",
    facts: [
      ['Société', company],
      ['Secteur', input.sector || '—'],
      ['Fondation', input.founded || '—'],
      ['Siège', input.headquarters || '—'],
      ['Effectifs', input.headcount || '—'],
      ['Actionnariat', input.ownership || '—'],
    ],
  });
  if (input.businessDescription || input.businessLines?.length) {
    slides.push({
      type: 'content',
      title: 'Activité',
      paragraphs: input.businessDescription ? [input.businessDescription] : [],
      bullets: input.businessLines || [],
    });
  }
  if (input.history?.length) {
    slides.push({ type: 'content', title: 'Historique et étapes clés', bullets: input.history });
  }
  if (input.products?.length) {
    slides.push({
      type: 'table',
      title: 'Offre produits / services',
      table: {
        headers: ['Offre', 'Description', 'Part du CA'],
        rows: input.products.map((p) => [p.name, p.description || '—', p.share || '—']),
      },
    });
  }

  // --- 4. Market & competition ---
  if (input.market) {
    const m = input.market;
    slides.push({ type: 'section', title: 'Marché et positionnement concurrentiel' });
    slides.push({
      type: 'content',
      title: 'Dynamique de marché',
      paragraphs: m.overview ? [m.overview] : [],
      bullets: m.drivers || [],
    });
    if (m.size?.values?.length) {
      slides.push({
        type: 'chart',
        title: `Taille du marché${m.size.unit ? ` (${m.size.unit})` : ''}`,
        chart: {
          kind: 'bar',
          series: [{ name: 'Marché', labels: (m.size.labels || []).map(String), values: m.size.values }],
        },
      });
    }
    if (m.competitors?.length) {
      slides.push({
        type: 'table',
        title: 'Paysage concurrentiel',
        table: {
          headers: ['Acteur', 'Positionnement', 'Part de marché'],
          rows: m.competitors.map((c) => [c.name, c.positioning || '—', c.share || '—']),
        },
      });
      const points = m.competitors
        .filter((c) => c.x != null && c.y != null)
        .map((c) => ({ name: c.name, x: c.x, y: c.y, highlight: c.name === company || c.highlight }));
      if (points.length >= 2) {
        slides.push({
          type: 'matrix',
          title: 'Matrice de positionnement concurrentiel',
          matrix: {
            xLabel: m.matrixXLabel || 'Étendue de l\'offre',
            yLabel: m.matrixYLabel || 'Qualité / valeur ajoutée',
            points,
          },
        });
      }
    }
  }

  // --- 5. Business model & clients ---
  if (input.revenueSplit?.values?.length || input.geographicSplit?.values?.length || input.clients) {
    slides.push({ type: 'section', title: 'Modèle économique et clients' });
    if (input.revenueSplit?.values?.length && input.geographicSplit?.values?.length) {
      slides.push({
        type: 'content',
        title: 'Répartition du chiffre d\'affaires',
        twoCol: {
          left: { chart: { kind: 'pie', title: input.revenueSplit.title || 'Par activité', labels: input.revenueSplit.labels, values: input.revenueSplit.values } },
          right: { chart: { kind: 'pie', title: 'Par géographie', labels: input.geographicSplit.labels, values: input.geographicSplit.values } },
        },
      });
    } else if (input.revenueSplit?.values?.length) {
      slides.push({
        type: 'chart',
        title: input.revenueSplit.title || 'Répartition du chiffre d\'affaires',
        chart: { kind: 'pie', labels: input.revenueSplit.labels, values: input.revenueSplit.values },
      });
    } else if (input.geographicSplit?.values?.length) {
      slides.push({
        type: 'chart',
        title: 'Répartition géographique du chiffre d\'affaires',
        chart: { kind: 'pie', labels: input.geographicSplit.labels, values: input.geographicSplit.values },
      });
    }
    if (input.clients?.concentration?.length) {
      slides.push({
        type: 'content',
        title: 'Base clients et concentration',
        paragraphs: input.clients.commentary ? [input.clients.commentary] : [],
        table: {
          headers: ['Client', 'Part du CA'],
          rows: input.clients.concentration.map((c) => [c.name, c.share || '—']),
        },
      });
    }
  }

  // --- 6. Management ---
  if (input.management?.length) {
    slides.push({ type: 'section', title: 'Équipe de management' });
    slides.push({
      type: 'table',
      title: 'Équipe dirigeante',
      table: {
        headers: ['Nom', 'Fonction', 'Expérience'],
        rows: input.management.map((mgr) => [mgr.name, mgr.role, mgr.bio || '—']),
      },
    });
  }

  // --- 7. Historical financials ---
  slides.push({ type: 'section', title: 'Performance financière historique' });
  if (fin.revenue?.length && years.length) {
    slides.push({
      type: 'chart',
      title: "Évolution du chiffre d'affaires et de l'EBITDA (m€)",
      chart: {
        kind: 'bar',
        series: [
          { name: 'CA', labels: years, values: fin.revenue },
          ...(fin.ebitda ? [{ name: 'EBITDA', labels: years, values: fin.ebitda }] : []),
        ],
      },
    });
    if (fin.ebitda?.length) {
      slides.push({
        type: 'chart',
        title: 'Marge EBITDA (%)',
        chart: {
          kind: 'line',
          series: [{
            name: 'Marge EBITDA',
            labels: years,
            values: fin.ebitda.map((e, i) => Math.round((e / fin.revenue[i]) * 1000) / 10),
          }],
        },
      });
    }
    const rows = [];
    if (fin.revenue) rows.push(["Chiffre d'affaires (m€)", ...fin.revenue]);
    if (fin.ebitda) rows.push(['EBITDA (m€)', ...fin.ebitda]);
    if (fin.ebitda && fin.revenue) rows.push(['Marge EBITDA', ...fin.ebitda.map((e, i) => formatPercent(e / fin.revenue[i]))]);
    if (fin.netIncome) rows.push(['Résultat net (m€)', ...fin.netIncome]);
    if (fin.capex) rows.push(['Capex (m€)', ...fin.capex]);
    if (fin.workingCapital) rows.push(['BFR (m€)', ...fin.workingCapital]);
    if (fin.netDebt) rows.push(['Dette nette (m€)', ...fin.netDebt]);
    slides.push({
      type: 'table',
      title: 'Synthèse financière',
      table: { headers: ['Agrégat', ...years], rows },
    });
  }

  // --- 8. Business plan ---
  if (bp.revenue?.length || bp.assumptions?.length || bp.commentary) {
    slides.push({ type: 'section', title: 'Business plan et perspectives' });
    if (bp.revenue?.length && bp.years?.length) {
      slides.push({
        type: 'chart',
        title: 'Projections financières (m€)',
        chart: {
          kind: 'bar',
          series: [
            { name: 'CA projeté', labels: bp.years.map(String), values: bp.revenue },
            ...(bp.ebitda ? [{ name: 'EBITDA projeté', labels: bp.years.map(String), values: bp.ebitda }] : []),
          ],
        },
      });
    }
    slides.push({
      type: 'content',
      title: 'Hypothèses du business plan',
      paragraphs: bp.commentary ? [bp.commentary] : [],
      bullets: bp.assumptions || [],
    });
  }

  // --- 9. Strategy ---
  if (input.strategy?.length) {
    slides.push({ type: 'content', title: 'Axes stratégiques de développement', bullets: input.strategy });
  }

  // --- 10. Risks ---
  if (input.risks?.length) {
    slides.push({
      type: 'table',
      title: 'Facteurs de risque et mitigants',
      table: {
        headers: ['Risque', 'Mitigant'],
        rows: input.risks.map((r) => (Array.isArray(r) ? r : [r.risk || r, r.mitigant || '—'])),
      },
    });
  }

  // --- 11. Transaction ---
  slides.push({ type: 'section', title: 'Transaction envisagée' });
  slides.push({
    type: 'facts',
    title: 'Cadre de la transaction',
    facts: [
      ['Nature', input.transaction?.type || 'Cession majoritaire'],
      ['Périmètre', input.transaction?.scope || '100% du capital'],
      ['Processus', input.transaction?.process || 'Processus compétitif'],
      ['Calendrier', input.transaction?.timeline || 'À déterminer'],
      ['Conseil', 'M&IA'],
    ],
  });

  // --- Custom additional sections / annexes ---
  if (input.sections?.length) {
    slides.push({ type: 'section', title: 'Annexes' });
    for (const s of input.sections) {
      slides.push({ type: 'content', title: s.title, paragraphs: s.paragraphs, bullets: s.bullets, table: s.table, chart: s.chart });
    }
  }

  return buildDeck({
    deckTitle: 'Confidential Information Memorandum',
    deckSubtitle: company,
    project: input.project,
    date: input.date,
    lang: 'fr',
    slides,
    disclaimer: disclaimerFor('ma', metadata.id),
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
