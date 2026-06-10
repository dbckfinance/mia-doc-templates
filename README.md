# mia-doc-templates

Moteur de génération déterministe de documents financiers M&IA : 65 générateurs couvrant M&A, Investment Banking, Audit, Hedge Fund et Equity Research.

## Principes

- **Déterministe** : même input JSON = même document. Aucun LLM dans la boucle de génération.
- **Toujours valide** : les scripts utilisent `docx`, `exceljs`, `pptxgenjs` — jamais du texte renommé en `.docx`. Chaque sortie est validée (magic bytes ZIP/`%PDF`).
- **Branding M&IA** : couleurs, polices, mentions de confidentialité et disclaimers centralisés dans `shared/`.
- **Conventions par vertical** : IB utilise « Exhibit », Audit « Annexe », ER « Figure ».

## Usage

```bash
npm install

# Lister les générateurs disponibles
npm run list
node registry.js --json        # catalogue machine-readable pour l'orchestrateur

# Générer un document (input JSON via fichier ou stdin)
node generators/ma/note-synthese.js tests/fixtures/note-synthese.json
cat input.json | node generators/ib/lbo-model.js
# → le fichier est écrit dans ./artifacts/

# Tests
npm run lint   # syntax check de tous les modules
npm test       # exécute chaque générateur et valide les sorties
```

## Structure

```
shared/          Moteur partagé : branding, layouts docx, tables, composeurs
                 déclaratifs (docx-report, pptx-deck, xlsx-model), charts,
                 helpers financiers, disclaimers, CLI standard
generators/
  ma/            14 générateurs M&A (note de synthèse, teaser, CIM, LOI, NDA…)
  ib/            12 générateurs IB (DCF, LBO, comps, football field…)
  audit/         12 générateurs Audit (rapport, QoE, VDD, risk assessment…)
  hedge-fund/    12 générateurs HF (investment memo, risk report, tearsheet…)
  equity-research/ 10 générateurs ER (initiation, earnings update, primer…)
  common/        5 générateurs transverses (exec summary, board memo, VDR index…)
tests/           Harness de validation + fixtures
```

## Interface standard d'un générateur

Chaque générateur exporte `metadata` (id, vertical, outputType, inputs requis) et `generate(input) → Buffer | [{ buffer, ext, suffix }]`, plus un mode CLI :

```bash
node generators/<vertical>/<id>.js input.json
```

Les sorties sont écrites dans `./artifacts/` (répertoire détecté par le Cursor SDK).

## Intégration Supercomputer

1. L'orchestrateur interroge `registry.js --json` pour mapper l'intention utilisateur vers un script.
2. Le Cursor Agent exécute `node generators/<vertical>/<id>.js` avec l'input JSON structuré plutôt que de générer le document de zéro.
3. Fallback : si aucun générateur ne correspond, génération libre avec les helpers `shared/` disponibles.
