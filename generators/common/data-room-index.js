// generators/common/data-room-index.js
// Index de data room / VDR (xlsx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildWorkbook } from '../../shared/xlsx-model.js';

export const metadata = {
  id: 'data-room-index',
  name: 'Index de Data Room (VDR)',
  vertical: 'common',
  outputType: 'xlsx',
  estimatedPages: '1-2 sheets',
  requiredInput: ['project'],
  optionalInput: ['sections', 'company'],
};

const DEFAULT_SECTIONS = [
  { code: '1', name: 'Corporate', items: ['Statuts et Kbis', 'Pactes d\'actionnaires', 'PV d\'assemblées et de conseils (3 ans)', 'Organigramme juridique'] },
  { code: '2', name: 'Finance', items: ['Comptes annuels audités (3 exercices)', 'Situations intermédiaires', 'Budget et business plan', 'Reporting mensuel de gestion', 'Détail de l\'endettement et tableaux d\'amortissement'] },
  { code: '3', name: 'Fiscal', items: ['Liasses fiscales (3 exercices)', 'Contrôles fiscaux et redressements', 'Documentation prix de transfert', 'Intégration fiscale'] },
  { code: '4', name: 'Juridique', items: ['Contrats clients significatifs', 'Contrats fournisseurs significatifs', 'Baux et titres de propriété', 'Litiges en cours et provisions', 'Propriété intellectuelle'] },
  { code: '5', name: 'Social / RH', items: ['Organigramme et effectifs', 'Contrats de travail des dirigeants et hommes clés', 'Accords collectifs', 'Plans d\'intéressement et de participation', 'Contentieux prud\'homaux'] },
  { code: '6', name: 'Opérations', items: ['Description des sites et capacités', 'Principaux équipements et capex', 'Certifications qualité', 'Assurances'] },
  { code: '7', name: 'IT', items: ['Cartographie du SI', 'Contrats de licences et TMA', 'Politique de cybersécurité', 'RGPD : registre des traitements'] },
  { code: '8', name: 'Commercial', items: ['Top 20 clients (CA, ancienneté, contrats)', 'Pipeline commercial', 'Études de marché', 'Politique tarifaire'] },
];

export async function generate(input = {}) {
  const project = input.project || 'Projet';
  const sections = input.sections?.length ? input.sections : DEFAULT_SECTIONS;

  const rows = [];
  for (const section of sections) {
    rows.push([`${section.code}.`, section.name.toUpperCase(), '', '', '']);
    (section.items || []).forEach((item, i) => {
      rows.push([`${section.code}.${i + 1}`, typeof item === 'string' ? item : item.name, statusOf(item), ownerOf(item), commentOf(item)]);
    });
  }

  return buildWorkbook({
    title: `Index VDR — ${project}`,
    sheets: [
      {
        name: 'Index VDR',
        sectionTitle: `Index de data room — ${project}${input.company ? ` (${input.company})` : ''}`,
        table: {
          headers: ['Réf.', 'Document', 'Statut', 'Responsable', 'Commentaire'],
          rows,
        },
        columns: [{ width: 8 }, { width: 56 }, { width: 14 }, { width: 16 }, { width: 30 }],
        freezeHeader: true,
      },
    ],
  });
}

function statusOf(item) { return typeof item === 'object' ? (item.status || 'À charger') : 'À charger'; }
function ownerOf(item) { return typeof item === 'object' ? (item.owner || '—') : '—'; }
function commentOf(item) { return typeof item === 'object' ? (item.comment || '') : ''; }

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
