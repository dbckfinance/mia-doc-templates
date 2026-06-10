// generators/hedge-fund/quarterly-letter.js
// Lettre trimestrielle aux investisseurs (docx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildLetter } from '../../shared/docx-report.js';
import { formatPercent } from '../../shared/financial-helpers.js';

export const metadata = {
  id: 'quarterly-letter',
  name: 'Lettre Trimestrielle aux Investisseurs',
  vertical: 'hedge-fund',
  outputType: 'docx',
  estimatedPages: '4-8',
  requiredInput: ['fund'],
  optionalInput: ['quarter', 'performance', 'commentary', 'topPositions', 'outlook', 'manager'],
};

export async function generate(input = {}) {
  const fund = input.fund || 'Fonds';
  const quarter = input.quarter || `T${Math.ceil((new Date().getMonth() + 1) / 3)} ${new Date().getFullYear()}`;
  const perf = input.performance || { quarter: 0.032, ytd: 0.078, benchmark: 0.021 };

  const paragraphs = [
    `Chers investisseurs,`,
    `Au titre du ${quarter}, ${fund} affiche une performance nette de ${formatPercent(perf.quarter)}, portant la performance depuis le début de l'année à ${formatPercent(perf.ytd)}, contre ${formatPercent(perf.benchmark)} pour notre indice de référence sur le trimestre.`,
    ...(input.commentary || [
      'Le trimestre a été marqué par une bonne tenue de nos positions de conviction, notamment dans les secteurs où notre travail fondamental nous donne un avantage informationnel durable. Notre discipline de construction de portefeuille — concentration sur nos meilleures idées, couverture systématique des risques de marché — a permis de capter l\'essentiel de la hausse tout en limitant la volatilité.',
      'Nous avons initié plusieurs nouvelles positions au cours de la période, dont le détail figure dans la section dédiée. À l\'inverse, nous avons soldé les positions dont la thèse était arrivée à maturité ou invalidée.',
    ]),
    ...(input.topPositions?.length
      ? [`Nos principales contributions du trimestre : ${input.topPositions.map((p) => `${p.name} (${formatPercent(p.contribution)})`).join(', ')}.`]
      : []),
    input.outlook || 'Pour les prochains trimestres, nous restons constructifs mais sélectifs. La dispersion des valorisations crée un environnement favorable à la sélection de titres, qui demeure le cœur de notre processus. Notre exposition nette reste modérée, et notre poche de liquidités nous permet de saisir les opportunités créées par la volatilité.',
    'Nous vous remercions de votre confiance.',
  ];

  return buildLetter({
    docTitle: `Lettre aux investisseurs — ${quarter}`,
    docSubtitle: fund,
    vertical: 'hedge-fund',
    docId: 'quarterly-letter',
    confidential: true,
    recipient: 'Aux investisseurs du fonds',
    paragraphs,
    signature: {
      name: input.manager || 'L\'équipe de gestion',
      title: fund,
    },
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
