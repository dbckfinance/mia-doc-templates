#!/usr/bin/env node
// Technofab CIM — Confidential Information Memorandum (50+ slides)
import fs from 'node:fs';
import path from 'node:path';
import { buildDeck } from '../shared/pptx-deck.js';
import { BRAND } from '../shared/branding.js';

const DISCLAIMER = `IMPORTANT NOTICE

This Confidential Information Memorandum (the "Memorandum") has been prepared by M&IA Intelligence ("M&IA") solely for informational purposes in connection with a potential transaction involving Technofab SAS and its subsidiaries (the "Company"). This Memorandum does not constitute an offer to sell or a solicitation of an offer to buy any securities or assets.

Forward-Looking Statements: This Memorandum contains forward-looking statements based on current expectations, estimates, forecasts, and projections about the Company and the industries in which it operates. These statements are subject to risks, uncertainties, and assumptions that could cause actual results to differ materially. No assurance can be given that future results will be achieved.

Confidentiality: This Memorandum is strictly confidential and is being furnished on a confidential basis to a limited number of qualified parties. Recipients must not reproduce, distribute, or disclose this Memorandum or any information contained herein to any third party without the prior written consent of M&IA and the Company. By accepting this Memorandum, you agree to return or destroy all copies upon request.

No Reliance: While the information contained herein has been prepared from sources believed to be reliable, neither M&IA nor the Company makes any representation or warranty, express or implied, as to the accuracy or completeness of the information. Recipients should conduct their own independent investigation and analysis. M&IA and the Company expressly disclaim any liability for any loss or damage arising from the use of this Memorandum or reliance on its contents.

No Contract: Nothing in this Memorandum shall form the basis of any contract or commitment. Any transaction will be subject to definitive agreements and customary closing conditions.`;

const YEARS_HIST = ['2022', '2023', '2024'];
const YEARS_PROJ = ['2025', '2026', '2027'];
const ALL_YEARS = [...YEARS_HIST, ...YEARS_PROJ];

const FIN = {
  revenue: [145, 160, 185, 205, 230, 260],
  ebitda: [23, 27, 37, 43, 51, 60],
  netIncome: [12, 15, 22, 26, 32, 40],
  ebitdaMargin: [15.9, 16.9, 20.0, 21.0, 22.2, 23.1],
};

const HIGHLIGHTS = [
  {
    title: 'European Leader in Industrial Automation Solutions',
    desc: 'Technofab has established a leading position in Western European industrial automation, combining hardware expertise with integrated software platforms. The Company benefits from strong brand recognition and deep customer relationships built over 25+ years.',
    bullets: [
      '18% market share in Western Europe industrial automation',
      '#1 position in France, #3 in Germany',
      '2,500+ active industrial clients across diverse end-markets',
      'Operations spanning 12 European countries',
    ],
  },
  {
    title: 'Recurring Revenue Model with High Visibility',
    desc: 'A significant portion of Technofab\'s revenue is recurring, driven by maintenance contracts and software subscriptions. This provides exceptional revenue visibility and supports predictable cash flow generation.',
    bullets: [
      '65% of revenue from recurring sources (maintenance + SaaS)',
      'Average contract duration of 4.2 years',
      'Net Revenue Retention of 112%',
      'Customer churn consistently below 3% annually',
    ],
  },
  {
    title: 'Proven Track Record of Profitable Growth',
    desc: 'Technofab has delivered consistent organic growth while expanding margins through operational leverage and mix shift toward higher-margin software and services.',
    bullets: [
      'Revenue CAGR of 12% (2021–2024)',
      'EBITDA margin expansion from 16% to 20%',
      'Three successful bolt-on acquisitions integrated',
      'Free cash flow conversion exceeding 70%',
    ],
  },
  {
    title: 'Strong Exposure to Secular Growth Trends',
    desc: 'The Company is well-positioned to benefit from structural tailwinds driving industrial automation adoption across European manufacturing.',
    bullets: [
      'Industry 4.0 and smart factory adoption (+12% annual investment growth)',
      'Reshoring and supply chain localization driving capex',
      'Energy efficiency mandates accelerating retrofit demand',
      'Labor shortages increasing automation ROI across sectors',
    ],
  },
  {
    title: 'Scalable Platform with Multiple Growth Levers',
    desc: 'Technofab\'s integrated platform offers multiple avenues for value creation beyond core market growth, including geographic expansion, product innovation, and strategic M&A.',
    bullets: [
      'Geographic expansion into UK, Nordics, and CEE (€500M+ TAM)',
      'New predictive maintenance product line launching H2 2025',
      'Cross-selling software to 2,500+ hardware customer base',
      'M&A pipeline with 15+ identified bolt-on targets',
    ],
  },
  {
    title: 'Experienced Management Team with Skin in the Game',
    desc: 'The leadership team combines deep industry expertise with a proven track record of value creation. Management is aligned with shareholders and committed to the transaction.',
    bullets: [
      'Average management tenure exceeding 8 years',
      'Management owns 15% of equity capital',
      'Successful integration of 3 acquisitions since 2018',
      'Committed to rolling over significant stake in transaction',
    ],
  },
];

function pct(v) { return `${v.toFixed(1)}%`; }

const slides = [
  // Disclaimer (slide 2)
  {
    type: 'content',
    title: 'Disclaimer',
    paragraphs: [DISCLAIMER],
  },

  // Table of Contents
  {
    type: 'content',
    title: 'Table of Contents',
    bullets: [
      'I. Executive Summary',
      'II. Investment Highlights',
      'III. Company Overview',
      'IV. Market Analysis',
      'V. Products & Services',
      'VI. Business Model',
      'VII. Operations Overview',
      'VIII. Management Team',
      'IX. Financial Overview',
      'X. Valuation Analysis',
      'XI. Risk Factors',
      'XII. Appendix',
    ],
  },

  // Executive Summary
  { type: 'section', title: 'Executive Summary' },
  {
    type: 'content',
    title: 'Transaction Overview',
    paragraphs: [
      'M&IA Intelligence has been retained by Technofab\'s shareholders to explore a strategic transaction. Technofab is a leading European provider of industrial automation solutions, offering an integrated portfolio of hardware, software, and services to manufacturing clients across 12 countries.',
      'The Company represents a compelling investment opportunity combining market leadership, recurring revenue visibility, and exposure to secular growth trends in industrial automation.',
    ],
    bullets: [
      'Leading European industrial automation platform with €185M revenue (2024)',
      '65% recurring revenue with 112% Net Revenue Retention',
      'EBITDA margin of 20% with strong FCF conversion (>70%)',
      'Multiple growth levers: geographic expansion, software cross-sell, M&A',
      'Experienced management team with 15% equity ownership',
      'Attractive valuation range of €350M–€480M (10.5x–13.0x EV/EBITDA)',
    ],
  },
  {
    type: 'table',
    title: 'Key Financial Metrics',
    table: {
      headers: ['Metric', 'Value', 'Commentary'],
      rows: [
        ['Revenue (2024)', '€185M', '12% CAGR (2021–2024)'],
        ['EBITDA (2024)', '€37M', '20.0% margin'],
        ['Net Debt', '€28M', '0.8x Net Debt / EBITDA'],
        ['Employees', '850 FTEs', 'Across 12 countries'],
        ['Recurring Revenue', '65%', 'Maintenance + software subscriptions'],
        ['Market Share (WE)', '18%', '#1 France, #3 Germany'],
      ],
    },
  },
  {
    type: 'content',
    title: 'Transaction Rationale',
    paragraphs: [
      'The shareholders are initiating a competitive sale process to identify a partner capable of supporting Technofab\'s next phase of growth. The transaction offers a unique opportunity to acquire a scaled, profitable platform in a fragmented market with significant consolidation potential.',
    ],
    bullets: [
      'Platform acquisition in a €45bn European TAM growing at 8% CAGR',
      'Opportunity to accelerate geographic expansion and product development',
      'Strong cross-selling potential between hardware and software offerings',
      'Proven M&A integration capabilities with active bolt-on pipeline',
      'Management committed to rollover and continued leadership',
    ],
  },

  // Investment Highlights
  { type: 'section', title: 'Investment Highlights' },
  ...HIGHLIGHTS.map((h) => ({
    type: 'content',
    title: h.title,
    paragraphs: [h.desc],
    bullets: h.bullets,
  })),

  // Company Overview
  { type: 'section', title: 'Company Overview' },
  {
    type: 'facts',
    title: 'Company Profile',
    facts: [
      ['Legal Name', 'Technofab SAS'],
      ['Founded', '1998 — Lyon, France'],
      ['Headquarters', 'Lyon, France'],
      ['Sector', 'Industrial Automation & Smart Manufacturing'],
      ['Employees', '850 FTEs'],
      ['Countries', '12 European countries'],
      ['Ownership', 'Founders & management (15%), PE fund (60%), Others (25%)'],
    ],
  },
  {
    type: 'content',
    title: 'Corporate History & Key Milestones',
    bullets: [
      { text: '1998 — Founded in Lyon as a PLC programming and integration specialist', level: 0 },
      { text: '2005 — Launched proprietary SCADA software platform (TechnoSCADA)', level: 0 },
      { text: '2012 — Expanded into Germany via acquisition of Automatik GmbH (Munich)', level: 0 },
      { text: '2018 — PE-backed buyout; initiated software-first strategy and MES platform', level: 0 },
      { text: '2022 — Acquired Polish robotics integrator RoboTech (Warsaw facility)', level: 0 },
      { text: '2024 — Revenue reaches €185M; predictive maintenance SaaS beta launch', level: 0 },
    ],
  },
  {
    type: 'content',
    title: 'Geographic Footprint',
    paragraphs: [
      'Technofab operates across 12 European countries with manufacturing facilities in Lyon (France), Munich (Germany), and Warsaw (Poland). R&D centers are located in Lyon and Grenoble.',
    ],
    bullets: [
      'France: HQ Lyon, 320 employees — largest market (38% of revenue)',
      'Germany: Munich facility, 210 employees — #3 market position',
      'Poland: Warsaw facility, 145 employees — cost-competitive manufacturing',
      'Benelux, Iberia, Italy, UK: Sales & service offices (175 employees)',
      'Nordics & CEE: Emerging markets targeted for expansion',
    ],
  },
  {
    type: 'chart',
    title: 'Revenue by Business Segment (2024)',
    chart: {
      kind: 'pie',
      labels: ['Hardware Solutions (45%)', 'Software Platform (30%)', 'Services (25%)'],
      values: [45, 30, 25],
      title: 'Segment Mix',
    },
  },

  // Market Analysis
  { type: 'section', title: 'Market Analysis' },
  {
    type: 'table',
    title: 'Market Size & Opportunity',
    table: {
      headers: ['Metric', 'Value', 'Definition'],
      rows: [
        ['TAM', '€45bn', 'Total European industrial automation market'],
        ['SAM', '€12bn', 'Addressable segments (mid-market manufacturing)'],
        ['SOM', '€2.2bn', 'Current serviceable market (Technofab segments)'],
        ['Market Growth', '8% CAGR', '2024–2028 projected growth rate'],
        ['Technofab Share', '18%', 'Western Europe market share'],
      ],
    },
  },
  {
    type: 'content',
    title: 'Key Market Trends',
    bullets: [
      'Industry 4.0 adoption: 68% of EU manufacturers investing in smart factory tech by 2027',
      'Labor shortage: 3.2M unfilled manufacturing jobs in EU driving automation demand',
      'Energy efficiency: EU Green Deal mandates 32.5% energy reduction by 2030',
      'Reshoring: 42% of EU manufacturers planning near-shoring initiatives (2024 survey)',
      'Predictive maintenance: Market growing at 15% CAGR, reaching €8bn by 2028',
      'Software penetration: Industrial software share rising from 22% to 35% of automation spend',
      'Consolidation: Fragmented market with 200+ mid-size players ripe for roll-up',
      'Digital twin adoption: 45% of large manufacturers deploying digital twin solutions',
    ],
  },
  {
    type: 'chart',
    title: 'Competitive Landscape — Market Share (Western Europe)',
    chart: {
      kind: 'bar',
      series: [{
        name: 'Market Share %',
        labels: ['Siemens', 'ABB', 'Technofab', 'Rockwell', 'Others'],
        values: [25, 22, 18, 15, 20],
      }],
      title: 'Industrial Automation Market Share',
    },
  },
  {
    type: 'table',
    title: 'Competitive Positioning',
    table: {
      headers: ['Competitor', 'Share', 'Strengths', 'Technofab Advantage'],
      rows: [
        ['Siemens', '25%', 'Scale, brand, full-stack', 'Agility, mid-market focus, local service'],
        ['ABB', '22%', 'Robotics leadership', 'Integrated software, lower TCO'],
        ['Rockwell', '15%', 'US market, PLC dominance', 'European footprint, open architecture'],
        ['Technofab', '18%', 'Integrated HW/SW/Services', '#1 France, recurring revenue model'],
      ],
    },
  },

  // Products & Services
  { type: 'section', title: 'Products & Services' },
  {
    type: 'content',
    title: 'Hardware Solutions (45% of Revenue)',
    paragraphs: [
      'Technofab designs and distributes industrial automation hardware including collaborative robots, PLCs, sensors, and motion control systems. Hardware serves as the entry point for long-term customer relationships.',
    ],
    bullets: [
      'Industrial robots: 6-axis, collaborative, and SCARA configurations',
      'PLCs & controllers: Proprietary TechnoPLC line + partner integrations',
      'Sensors & IoT: Temperature, pressure, vibration, vision systems',
      'Motion control: Servo drives, linear actuators, gantry systems',
    ],
  },
  {
    type: 'content',
    title: 'Software Platform (30% of Revenue)',
    paragraphs: [
      'Technofab\'s software suite provides end-to-end manufacturing intelligence, from shop-floor monitoring to predictive analytics. Software drives recurring revenue and customer stickiness.',
    ],
    bullets: [
      'TechnoSCADA: Real-time monitoring and control (1,200+ deployments)',
      'TechnoMES: Manufacturing execution system with OEE tracking',
      'TechnoPredict: AI-powered predictive maintenance (beta, GA H2 2025)',
      'TechnoConnect: Cloud-based integration hub and API platform',
    ],
  },
  {
    type: 'content',
    title: 'Services (25% of Revenue)',
    paragraphs: [
      'High-touch services complement hardware and software offerings, ensuring successful deployment and long-term customer satisfaction.',
    ],
    bullets: [
      'System integration: Turnkey automation project delivery',
      'Maintenance contracts: 4.2-year average duration, 65% renewal rate',
      'Training & certification: Technofab Academy (2,500+ certified engineers)',
      'Consulting: Digital transformation and Industry 4.0 roadmaps',
    ],
  },

  // Business Model
  { type: 'section', title: 'Business Model' },
  {
    type: 'content',
    title: 'Revenue Streams & Customer Segmentation',
    twoCol: {
      left: {
        chart: {
          kind: 'pie',
          labels: ['Hardware (45%)', 'Software (30%)', 'Services (25%)'],
          values: [45, 30, 25],
          title: 'Revenue Mix',
        },
      },
      right: {
        chart: {
          kind: 'pie',
          labels: ['Automotive (35%)', 'Food & Bev (25%)', 'Pharma (20%)', 'Other (20%)'],
          values: [35, 25, 20, 20],
          title: 'Customer Segments',
        },
      },
    },
  },
  {
    type: 'table',
    title: 'Key Business KPIs',
    table: {
      headers: ['KPI', '2022', '2023', '2024', 'Target 2025'],
      rows: [
        ['Recurring Revenue %', '58%', '62%', '65%', '68%'],
        ['Net Revenue Retention', '108%', '110%', '112%', '115%'],
        ['Customer Churn', '4.2%', '3.5%', '2.8%', '<3.0%'],
        ['Avg Contract Duration (yrs)', '3.8', '4.0', '4.2', '4.5'],
        ['Gross Margin', '42%', '44%', '46%', '48%'],
        ['Customer Satisfaction (NPS)', '62', '68', '74', '75+'],
      ],
    },
  },

  // Operations
  { type: 'section', title: 'Operations Overview' },
  {
    type: 'content',
    title: 'Manufacturing & R&D Facilities',
    bullets: [
      'Lyon, France: HQ, primary manufacturing (45,000 sqm), main R&D center',
      'Munich, Germany: Manufacturing & integration hub (28,000 sqm)',
      'Warsaw, Poland: Robotics assembly & cost-competitive production (22,000 sqm)',
      'Grenoble, France: Advanced R&D center — AI/ML and predictive analytics',
      'Total production capacity utilization: 78% (headroom for growth)',
    ],
  },
  {
    type: 'content',
    title: 'Operational Highlights',
    bullets: [
      'ISO 9001:2015, ISO 14001:2015, ISO 45001:2018 certified across all facilities',
      'Strategic partnerships: Siemens (components), Microsoft Azure (cloud), SAP (ERP integration)',
      'Supply chain: 85% of components sourced within EU; dual-sourcing on critical items',
      'Quality metrics: 99.2% on-time delivery, <0.5% defect rate',
      'R&D investment: €12M annually (6.5% of revenue), 120 dedicated engineers',
      'Patent portfolio: 34 active patents in automation and predictive maintenance',
      'Technofab Academy: Internal training center certifying 500+ engineers annually',
      'ESG: Carbon-neutral operations target by 2028; 35% renewable energy today',
    ],
  },

  // Management
  { type: 'section', title: 'Management Team' },
  {
    type: 'content',
    title: 'Organizational Structure',
    paragraphs: [
      'Technofab operates with a flat, matrix organization reporting to CEO Pierre Durand. Functional leaders oversee Engineering, Sales, Operations, and Finance across geographic regions.',
    ],
    bullets: [
      'CEO — Pierre Durand: Overall strategy, M&A, investor relations',
      'CFO — Marie Lambert: Finance, treasury, M&A integration',
      'COO — Hans Mueller: Manufacturing, supply chain, quality',
      'CTO — Sophie Martin: R&D, product development, IP',
      'CCO — Marco Rossi: Sales, marketing, customer success',
      'Regional VPs: France, DACH, CEE & Emerging Markets',
    ],
  },
  {
    type: 'table',
    title: 'Key Executives',
    table: {
      headers: ['Name', 'Role', 'Tenure', 'Background'],
      rows: [
        ['Pierre Durand', 'CEO', '15 years', 'Ex-Schneider Electric; INSEAD MBA'],
        ['Marie Lambert', 'CFO', '8 years', 'Ex-KPMG Transaction Services; HEC Paris'],
        ['Hans Mueller', 'COO', '10 years', 'Ex-Siemens Digital Industries; TU Munich'],
        ['Sophie Martin', 'CTO', '6 years', 'Ex-Dassault Systèmes; École Polytechnique'],
        ['Marco Rossi', 'CCO', '5 years', 'Ex-ABB Robotics; Bocconi University'],
      ],
    },
  },

  // Financial Overview
  { type: 'section', title: 'Financial Overview' },
  {
    type: 'chart',
    title: 'Revenue & EBITDA Evolution (€M)',
    chart: {
      kind: 'bar',
      series: [
        { name: 'Revenue', labels: ALL_YEARS, values: FIN.revenue },
        { name: 'EBITDA', labels: ALL_YEARS, values: FIN.ebitda },
      ],
      title: 'Historical & Projected Performance',
    },
  },
  {
    type: 'table',
    title: 'Profit & Loss Summary (€M)',
    table: {
      headers: ['', ...ALL_YEARS],
      rows: [
        ['Revenue', ...FIN.revenue],
        ['EBITDA', ...FIN.ebitda],
        ['EBITDA Margin', ...FIN.ebitdaMargin.map(pct)],
        ['Net Income', ...FIN.netIncome],
      ],
    },
  },
  {
    type: 'chart',
    title: 'Revenue Bridge 2023 → 2024 (€M)',
    chart: {
      kind: 'waterfall',
      steps: [
        { label: '2023 Revenue', value: 160, isTotal: true },
        { label: 'Organic Growth', value: 18 },
        { label: 'Price/Mix', value: 5 },
        { label: 'M&A (RoboTech)', value: 4 },
        { label: 'FX Impact', value: -2 },
        { label: '2024 Revenue', value: 185, isTotal: true },
      ],
      title: 'Revenue Bridge',
    },
  },
  {
    type: 'chart',
    title: 'EBITDA Margin Evolution',
    chart: {
      kind: 'line',
      series: [{
        name: 'EBITDA Margin %',
        labels: ALL_YEARS,
        values: FIN.ebitdaMargin,
      }],
      title: 'Margin Expansion Trajectory',
    },
  },
  {
    type: 'table',
    title: 'Balance Sheet Summary (€M)',
    table: {
      headers: ['', '2022', '2023', '2024'],
      rows: [
        ['Total Assets', 142, 158, 175],
        ['Fixed Assets', 68, 72, 78],
        ['Net Working Capital', 38, 42, 45],
        ['Total Debt', 42, 38, 35],
        ['Cash & Equivalents', 18, 15, 7],
        ['Net Debt', 24, 23, 28],
        ['Shareholders\' Equity', 86, 98, 112],
      ],
    },
  },
  {
    type: 'table',
    title: 'Cash Flow Summary (€M)',
    table: {
      headers: ['', '2022', '2023', '2024'],
      rows: [
        ['EBITDA', 23, 27, 37],
        ['Capex', -8, -9, -11],
        ['Change in NWC', -3, -4, -3],
        ['Operating Cash Flow', 12, 14, 23],
        ['Free Cash Flow', 9, 11, 18],
        ['FCF Conversion', '39%', '41%', '49%'],
        ['Interest & Taxes', -5, -6, -8],
        ['M&A Outflows', 0, -8, -4],
      ],
    },
  },
  {
    type: 'table',
    title: 'Key Financial KPIs',
    table: {
      headers: ['Metric', '2022', '2023', '2024'],
      rows: [
        ['Revenue Growth', '—', '10.3%', '15.6%'],
        ['Gross Margin', '42%', '44%', '46%'],
        ['EBITDA Margin', '15.9%', '16.9%', '20.0%'],
        ['Net Debt / EBITDA', '1.0x', '0.9x', '0.8x'],
        ['ROCE', '14%', '16%', '19%'],
        ['Days Sales Outstanding', 58, 55, 52],
        ['Inventory Turnover', 4.2, 4.5, 4.8],
      ],
    },
  },
  {
    type: 'content',
    title: 'Financial Outlook & Assumptions',
    paragraphs: [
      'Management projections for 2025–2027 assume continued organic growth of 10–12% annually, driven by software mix shift, geographic expansion, and new product launches. EBITDA margins are expected to expand to 23% by 2027 through operational leverage.',
    ],
    bullets: [
      '2025E Revenue: €205M (+11%); EBITDA: €43M (21.0% margin)',
      '2026E Revenue: €230M (+12%); EBITDA: €51M (22.2% margin)',
      '2027E Revenue: €260M (+13%); EBITDA: €60M (23.1% margin)',
      'Capex maintained at 5–6% of revenue; working capital stable',
    ],
  },

  // Valuation
  { type: 'section', title: 'Valuation Analysis' },
  {
    type: 'chart',
    title: 'Valuation Football Field (Enterprise Value, €M)',
    chart: {
      kind: 'footballField',
      ranges: [
        { method: 'DCF Analysis', low: 380, high: 450 },
        { method: 'Trading Comps', low: 350, high: 420 },
        { method: 'Transaction Comps', low: 400, high: 480 },
        { method: 'LBO Analysis', low: 360, high: 430 },
      ],
      title: 'Enterprise Value Ranges (€M)',
      currency: '€',
    },
  },
  {
    type: 'table',
    title: 'Implied Valuation Multiples',
    table: {
      headers: ['Methodology', 'Low EV', 'Mid EV', 'High EV', 'EV/EBITDA (2024)'],
      rows: [
        ['DCF', '€380M', '€415M', '€450M', '10.3x – 12.2x'],
        ['Trading Comps', '€350M', '€385M', '€420M', '9.5x – 11.4x'],
        ['Transaction Comps', '€400M', '€440M', '€480M', '10.8x – 13.0x'],
        ['LBO', '€360M', '€395M', '€430M', '9.7x – 11.6x'],
        ['Consolidated Range', '€350M', '€410M', '€480M', '10.5x – 13.0x'],
      ],
    },
  },
  {
    type: 'table',
    title: 'Comparable Companies',
    table: {
      headers: ['Company', 'EV (€M)', 'Revenue (€M)', 'EBITDA Margin', 'EV/EBITDA'],
      rows: [
        ['Siemens Digital Industries', '12,500', '4,800', '18%', '12.5x'],
        ['ABB Robotics & Discrete', '8,200', '3,100', '17%', '11.8x'],
        ['Rockwell Automation', '6,800', '2,400', '22%', '13.2x'],
        ['Schneider Electric (IA)', '4,500', '1,800', '19%', '12.0x'],
        ['KUKA AG', '2,100', '950', '15%', '10.5x'],
        ['Comau (Stellantis)', '1,400', '620', '14%', '9.8x'],
        ['Peer Median', '—', '—', '17%', '11.5x'],
        ['Technofab (Implied)', '410', '185', '20%', '11.1x'],
      ],
    },
  },

  // Risk Factors
  { type: 'section', title: 'Risk Factors' },
  {
    type: 'content',
    title: 'Key Risk Factors',
    bullets: [
      { text: 'Market Risks', level: 0 },
      { text: 'Competition from larger players (Siemens, ABB) with greater R&D budgets', level: 1 },
      { text: 'Cyclicality in automotive and manufacturing capex cycles', level: 1 },
      { text: 'Technology disruption from AI-native automation startups', level: 1 },
      { text: 'Operational Risks', level: 0 },
      { text: 'Supply chain disruption for semiconductor and electronic components', level: 1 },
      { text: 'Key personnel dependency on senior management team', level: 1 },
      { text: 'Integration risk from ongoing M&A activity', level: 1 },
      { text: 'Financial Risks', level: 0 },
      { text: 'FX exposure: 35% of revenue in non-EUR currencies', level: 1 },
      { text: 'Customer concentration: Top 10 clients represent 28% of revenue', level: 1 },
      { text: 'Regulatory Risks', level: 0 },
      { text: 'Evolving EU machinery directive and cybersecurity requirements', level: 1 },
      { text: 'Data privacy compliance (GDPR) for cloud-based software platform', level: 1 },
    ],
  },
  {
    type: 'chart',
    title: 'Risk Assessment Matrix',
    chart: {
      kind: 'heatmap',
      risks: [
        { id: '1', impact: 4, likelihood: 3 },
        { id: '2', impact: 3, likelihood: 4 },
        { id: '3', impact: 5, likelihood: 2 },
        { id: '4', impact: 4, likelihood: 4 },
        { id: '5', impact: 3, likelihood: 3 },
        { id: '6', impact: 2, likelihood: 4 },
        { id: '7', impact: 3, likelihood: 2 },
        { id: '8', impact: 4, likelihood: 3 },
      ],
    },
  },

  // Appendix
  { type: 'section', title: 'Appendix' },
  {
    type: 'table',
    title: 'Appendix A — Detailed Income Statement (€M)',
    table: {
      headers: ['', '2022', '2023', '2024'],
      rows: [
        ['Revenue', 145, 160, 185],
        ['Cost of Goods Sold', -84, -90, -100],
        ['Gross Profit', 61, 70, 85],
        ['SG&A', -32, -35, -40],
        ['R&D', -8, -9, -10],
        ['Other Operating', -2, -2, -2],
        ['EBITDA', 23, 27, 37],
        ['D&A', -6, -7, -8],
        ['EBIT', 17, 20, 29],
        ['Net Financial Result', -2, -2, -3],
        ['Tax', -3, -3, -4],
        ['Net Income', 12, 15, 22],
      ],
    },
  },
  {
    type: 'table',
    title: 'Appendix B — Top 20 Customers (Anonymized)',
    table: {
      headers: ['#', 'Sector', 'Country', 'Revenue (€M)', '% of Total'],
      rows: [
        ['1', 'Automotive OEM', 'France', 8.2, '4.4%'],
        ['2', 'Food & Beverage', 'Germany', 6.5, '3.5%'],
        ['3', 'Pharmaceutical', 'Switzerland', 5.8, '3.1%'],
        ['4', 'Automotive Tier-1', 'France', 5.2, '2.8%'],
        ['5', 'Chemical', 'Germany', 4.8, '2.6%'],
        ['6', 'Food & Beverage', 'Italy', 4.2, '2.3%'],
        ['7', 'Aerospace', 'France', 3.9, '2.1%'],
        ['8', 'Automotive OEM', 'Spain', 3.5, '1.9%'],
        ['9', 'Pharmaceutical', 'Germany', 3.2, '1.7%'],
        ['10', 'Logistics', 'Netherlands', 2.9, '1.6%'],
        ['11–20', 'Various', 'Multi', 18.4, '9.9%'],
        ['Total Top 20', '', '', 66.6, '36.0%'],
      ],
    },
  },
  {
    type: 'table',
    title: 'Appendix C — Product Specifications Summary',
    table: {
      headers: ['Product Line', 'Key Specs', 'Deployments', 'Avg Deal Size'],
      rows: [
        ['TechnoRobot 6-Axis', 'Payload 3–50kg, reach 850–2100mm', '420 units/yr', '€85K'],
        ['TechnoPLC Series', '32–2048 I/O, SIL2 certified', '2,800 units/yr', '€12K'],
        ['TechnoSCADA', 'Cloud/on-prem, 50K tags, real-time', '1,200 licenses', '€45K/yr'],
        ['TechnoMES', 'OEE, traceability, batch management', '680 licenses', '€38K/yr'],
        ['TechnoPredict', 'ML-based, 48hr failure prediction', 'Beta (85 pilots)', '€25K/yr'],
        ['Maintenance Contracts', '24/7 support, SLA 4hr response', '1,800 contracts', '€18K/yr'],
      ],
    },
  },
  {
    type: 'content',
    title: 'Appendix D — Market Study Sources',
    bullets: [
      'MarketsandMarkets — European Industrial Automation Market Report 2024',
      'McKinsey & Company — Industry 4.0: Capturing Value in Manufacturing (2024)',
      'Statista — Industrial Robot Installations in Europe (2024)',
      'EUROSTAT — Manufacturing Employment & Automation Trends (2024)',
      'Gartner — Magic Quadrant for Manufacturing Execution Systems (2024)',
      'IDC — European IoT in Manufacturing Forecast 2024–2028',
      'Technofab Internal Market Intelligence — Competitive Benchmarking Study (2024)',
      'M&IA Intelligence — Sector Analysis & Valuation Benchmarking (2025)',
    ],
  },
  {
    type: 'content',
    title: 'Appendix E — Glossary of Terms',
    bullets: [
      'CAGR: Compound Annual Growth Rate',
      'EBITDA: Earnings Before Interest, Taxes, Depreciation & Amortization',
      'EV: Enterprise Value (Equity Value + Net Debt)',
      'FCF: Free Cash Flow (Operating Cash Flow less Capex)',
      'MES: Manufacturing Execution System',
      'NRR: Net Revenue Retention (expansion minus churn)',
      'OEE: Overall Equipment Effectiveness',
      'PLC: Programmable Logic Controller',
      'SCADA: Supervisory Control and Data Acquisition',
      'TAM/SAM/SOM: Total/Serviceable/Obtainable Addressable Market',
    ],
  },
];

async function main() {
  const artifactsDir = path.resolve('./artifacts');
  if (!fs.existsSync(artifactsDir)) fs.mkdirSync(artifactsDir, { recursive: true });

  const buffer = await buildDeck({
    deckTitle: 'Confidential Information Memorandum',
    deckSubtitle: 'Technofab',
    advisorName: 'M&IA Intelligence',
    project: 'Financial Advisor',
    date: 'June 2025',
    lang: 'en',
    confidential: true,
    slides,
    disclaimer: false,
  });

  const outPath = path.join(artifactsDir, 'Technofab_CIM_June2025.pptx');
  fs.writeFileSync(outPath, buffer);
  console.log(`OK: wrote ${outPath} (${buffer.length} bytes, ${slides.length + 1} slides)`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
