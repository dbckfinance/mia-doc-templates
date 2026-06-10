// generators/audit/tax-dd-report.js
// Rapport de due diligence fiscale (docx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildReport } from '../../shared/docx-report.js';

export const metadata = {
  id: 'tax-dd-report',
  name: 'Due Diligence Fiscale',
  vertical: 'audit',
  outputType: 'docx',
  estimatedPages: '10-20',
  requiredInput: ['company'],
  optionalInput: ['jurisdictions', 'findings', 'period', 'exposures'],
};

const DEFAULT_FINDINGS = [
  { area: 'Impôt sur les sociétés', risk: 'Moyen', detail: 'Déficits reportables dont l\'utilisation pourrait être remise en cause en cas de changement d\'activité.', exposure: '0,8 m€' },
  { area: 'TVA', risk: 'Faible', detail: 'Régularisations mineures identifiées sur la TVA déductible.', exposure: '< 0,1 m€' },
  { area: 'Prix de transfert', risk: 'Élevé', detail: 'Documentation de prix de transfert incomplète sur les flux intra-groupe de management fees.', exposure: '1,2 m€' },
  { area: 'Taxes locales', risk: 'Faible', detail: 'CFE/CVAE correctement déclarées sur la période.', exposure: '—' },
];

export async function generate(input = {}) {
  const company = input.company || 'Société';
  const findings = input.findings?.length ? input.findings : DEFAULT_FINDINGS;

  return buildReport({
    docTitle: 'Rapport de due diligence fiscale',
    docSubtitle: company,
    vertical: 'audit',
    docId: 'tax-dd-report',
    confidential: true,
    sections: [
      {
        heading: 'Synthèse',
        blocks: [
          { type: 'p', text: `Nous avons procédé à une revue fiscale de ${company} portant sur ${input.period || 'les trois derniers exercices non prescrits'}. Les principales zones de risque et expositions estimées sont synthétisées ci-dessous.` },
          {
            type: 'table',
            headers: ['Domaine', 'Niveau de risque', 'Exposition estimée'],
            rows: findings.map((f) => [f.area, f.risk, f.exposure || '—']),
          },
        ],
      },
      {
        heading: 'Périmètre',
        blocks: [
          { type: 'kv', label: 'Société', value: company },
          { type: 'kv', label: 'Juridictions couvertes', value: (input.jurisdictions || ['France']).join(', ') },
          { type: 'kv', label: 'Période', value: input.period || 'Exercices non prescrits' },
        ],
      },
      ...findings.map((f) => ({
        heading: `${f.area} — risque ${f.risk}`,
        blocks: [
          { type: 'p', text: f.detail },
          { type: 'kv', label: 'Exposition estimée', value: f.exposure || 'Non quantifiable à ce stade' },
          { type: 'kv', label: 'Recommandation', value: f.recommendation || 'Couverture via garantie de passif spécifique ou ajustement de prix.' },
        ],
      })),
      {
        heading: 'Recommandations contractuelles',
        blocks: [
          { type: 'bullets', items: [
            'Inclure une garantie de passif fiscale couvrant les exercices non prescrits',
            'Prévoir des garanties spécifiques sur les zones de risque élevé identifiées',
            'Obtenir les justificatifs des déficits reportables avant le closing',
          ] },
        ],
      },
      {
        heading: 'Limites',
        blocks: [
          { type: 'p', text: 'Cette revue ne constitue pas un audit fiscal exhaustif. Elle repose sur les documents communiqués en data room et les échanges avec la direction. Les expositions chiffrées sont des estimations en principal, hors intérêts de retard et pénalités sauf mention contraire.' },
        ],
      },
    ],
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
