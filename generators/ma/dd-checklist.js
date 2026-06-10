// generators/ma/dd-checklist.js
// Due diligence checklist par workstream (xlsx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildWorkbook } from '../../shared/xlsx-model.js';

export const metadata = {
  id: 'dd-checklist',
  name: 'Due Diligence Checklist',
  vertical: 'ma',
  outputType: 'xlsx',
  estimatedPages: '5-8 sheets',
  requiredInput: [],
  optionalInput: ['project', 'workstreams', 'customItems'],
};

const DEFAULT_WORKSTREAMS = {
  'Financière': [
    'Comptes annuels audités (3 derniers exercices)',
    'Situations intermédiaires et reporting mensuel',
    'Détail du chiffre d\'affaires par client / produit / géographie',
    'Analyse de la marge brute et pont EBITDA',
    'Normalisation de l\'EBITDA (éléments non récurrents)',
    'BFR normatif et saisonnalité',
    'Dette nette et éléments assimilés (provisions, engagements)',
    'Plan d\'affaires et hypothèses sous-jacentes',
    'Capex historiques et prévisionnels',
  ],
  'Juridique': [
    'Statuts et pactes d\'actionnaires',
    'Organigramme juridique du groupe',
    'Contrats clients et fournisseurs significatifs',
    'Contrats de financement et sûretés',
    'Litiges en cours et provisions associées',
    'Propriété intellectuelle (marques, brevets, licences)',
    'Conformité RGPD',
    'Baux commerciaux et actifs immobiliers',
  ],
  'Fiscale': [
    'Liasses fiscales (3 derniers exercices)',
    'Contrôles fiscaux passés et en cours',
    'Intégration fiscale et conventions intragroupe',
    'Prix de transfert (documentation)',
    'Crédits d\'impôt (CIR, CII) et subventions',
    'TVA et taxes assises sur les salaires',
  ],
  'Sociale': [
    'Effectifs et organigramme fonctionnel',
    'Contrats de travail des hommes clés',
    'Accords collectifs et usages',
    'Rémunérations variables et plans d\'intéressement',
    'Contentieux prud\'homaux',
    'Conformité durée du travail',
  ],
  'Opérationnelle / IT': [
    'Cartographie des systèmes d\'information',
    'Contrats de licence et maintenance IT',
    'Cybersécurité et plan de continuité',
    'Supply chain et dépendances fournisseurs',
    'Certifications qualité',
  ],
  'ESG': [
    'Politique environnementale et certifications',
    'Bilan carbone et plan de réduction',
    'Politique sociale et gouvernance',
    'Risques de conformité (anticorruption, sanctions)',
  ],
};

export async function generate(input = {}) {
  const workstreams = input.workstreams || DEFAULT_WORKSTREAMS;
  const sheets = [];

  const summaryRows = Object.entries(workstreams).map(([name, items]) => [name, items.length, 0, items.length]);
  sheets.push({
    name: 'Synthèse',
    sectionTitle: `Due Diligence Checklist — ${input.project || 'Projet'}`,
    table: {
      headers: ['Workstream', 'Items', 'Reçus', 'En attente'],
      rows: summaryRows,
    },
    columns: [{ width: 30 }, { width: 12 }, { width: 12 }, { width: 14 }],
  });

  for (const [name, items] of Object.entries(workstreams)) {
    sheets.push({
      name: name.slice(0, 31).replace(/[\\/?*[\]]/g, '-'),
      sectionTitle: `DD ${name}`,
      table: {
        headers: ['#', 'Document / Information demandée', 'Priorité', 'Statut', 'Référence VDR', 'Commentaires'],
        rows: items.map((item, i) => [
          `${name.slice(0, 3).toUpperCase()}-${String(i + 1).padStart(2, '0')}`,
          typeof item === 'string' ? item : item.label,
          typeof item === 'object' ? (item.priority || 'Normale') : 'Normale',
          'En attente',
          '',
          '',
        ]),
      },
      columns: [{ width: 10 }, { width: 60 }, { width: 12 }, { width: 14 }, { width: 16 }, { width: 30 }],
      freezeHeader: true,
    });
  }

  for (const custom of input.customItems || []) {
    // Allow appending custom items as their own sheet
    sheets.push({
      name: (custom.workstream || 'Divers').slice(0, 31),
      sectionTitle: `DD ${custom.workstream || 'Divers'}`,
      table: {
        headers: ['#', 'Document / Information demandée', 'Priorité', 'Statut', 'Référence VDR', 'Commentaires'],
        rows: (custom.items || []).map((item, i) => [i + 1, item, 'Normale', 'En attente', '', '']),
      },
      freezeHeader: true,
    });
  }

  return buildWorkbook({ title: `DD Checklist — ${input.project || 'Projet'}`, sheets });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
