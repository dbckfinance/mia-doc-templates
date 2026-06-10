// generators/audit/controls-testing.js
// Feuille de travail — tests de contrôles (xlsx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildWorkbook } from '../../shared/xlsx-model.js';

export const metadata = {
  id: 'controls-testing',
  name: 'Tests de Contrôles (Workpaper)',
  vertical: 'audit',
  outputType: 'xlsx',
  estimatedPages: '2-3 sheets',
  requiredInput: ['company'],
  optionalInput: ['cycle', 'controls', 'period', 'preparedBy'],
};

const DEFAULT_CONTROLS = [
  { id: 'C-01', description: 'Validation des factures fournisseurs > 5 k€ par le DAF', frequency: 'À chaque opération', samples: 25, exceptions: 0, conclusion: 'Efficace' },
  { id: 'C-02', description: 'Rapprochement bancaire mensuel formalisé et revu', frequency: 'Mensuel', samples: 12, exceptions: 1, conclusion: 'Efficace avec exception isolée' },
  { id: 'C-03', description: 'Revue des accès utilisateurs ERP', frequency: 'Trimestriel', samples: 4, exceptions: 2, conclusion: 'Inefficace — déficience' },
  { id: 'C-04', description: 'Approbation des notes de frais par le manager', frequency: 'À chaque opération', samples: 25, exceptions: 0, conclusion: 'Efficace' },
];

export async function generate(input = {}) {
  const company = input.company || 'Société';
  const cycle = input.cycle || 'Achats / Trésorerie / IT';
  const controls = input.controls?.length ? input.controls : DEFAULT_CONTROLS;

  const deficient = controls.filter((c) => /inefficace|déficien/i.test(c.conclusion || ''));

  return buildWorkbook({
    title: `Tests de contrôles — ${company}`,
    sheets: [
      {
        name: 'En-tête',
        sectionTitle: 'Feuille de travail — Tests de contrôles',
        table: {
          headers: ['Champ', 'Valeur'],
          rows: [
            ['Entité', company],
            ['Cycle(s)', cycle],
            ['Période', input.period || 'Exercice en cours'],
            ['Préparé par', input.preparedBy || 'Équipe audit M&IA'],
            ['Date', new Date().toLocaleDateString('fr-FR')],
            ['Nombre de contrôles testés', controls.length],
            ['Déficiences identifiées', deficient.length],
          ],
        },
        columns: [{ width: 28 }, { width: 40 }],
      },
      {
        name: 'Tests',
        sectionTitle: 'Détail des tests',
        table: {
          headers: ['Réf.', 'Description du contrôle', 'Fréquence', 'Échantillon', 'Exceptions', 'Conclusion'],
          rows: controls.map((c) => [c.id, c.description, c.frequency || '—', c.samples ?? '—', c.exceptions ?? '—', c.conclusion || 'À conclure']),
        },
        columns: [{ width: 8 }, { width: 52 }, { width: 18 }, { width: 12 }, { width: 12 }, { width: 30 }],
        freezeHeader: true,
      },
      {
        name: 'Déficiences',
        sectionTitle: 'Déficiences et impact sur la stratégie d\'audit',
        table: {
          headers: ['Réf.', 'Contrôle déficient', 'Impact'],
          rows: deficient.length
            ? deficient.map((c) => [c.id, c.description, 'Augmenter les tests substantifs sur l\'assertion concernée'])
            : [['—', 'Aucune déficience identifiée', 'Approche d\'audit inchangée']],
        },
        columns: [{ width: 8 }, { width: 52 }, { width: 44 }],
      },
    ],
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
