# Règles de génération de documents M&IA

## Règle n°1 : toujours utiliser un générateur existant

Avant de générer un document de zéro, consulte le catalogue :

```bash
node registry.js --json
```

Si un générateur correspond au type de document demandé, utilise-le :

```bash
# 1. Écrire l'input structuré
#    (champs requis dans metadata.requiredInput de chaque générateur)
echo '{ "company": "...", ... }' > /tmp/input.json

# 2. Exécuter le générateur — la sortie est écrite dans ./artifacts/
node generators/<vertical>/<id>.js /tmp/input.json
```

## Règle n°2 : ne JAMAIS renommer du texte en .docx/.xlsx/.pptx

Les fichiers Office sont des conteneurs ZIP. Écrire du markdown ou du texte
brut et le renommer en `.docx` produit un fichier corrompu. Utilise toujours
les bibliothèques (`docx`, `exceljs`, `pptxgenjs`) via les générateurs ou les
composeurs `shared/` (`docx-report.js`, `xlsx-model.js`, `pptx-deck.js`).

## Règle n°3 : sorties dans ./artifacts/

Tous les fichiers livrables doivent être écrits dans `./artifacts/` — c'est le
seul répertoire détecté pour le téléchargement.

## Règle n°4 : fallback en génération libre

Si aucun générateur ne correspond, compose le document avec les composeurs
déclaratifs de `shared/` qui appliquent automatiquement le branding M&IA :

```javascript
import { buildReport } from './shared/docx-report.js';   // docx
import { buildDeck } from './shared/pptx-deck.js';        // pptx
import { buildWorkbook } from './shared/xlsx-model.js';   // xlsx
```

## Correspondance verticaux

| Vertical | Répertoire | Convention exhibits |
|---|---|---|
| M&A / Advisory | `generators/ma/` | Annexe |
| Investment Banking | `generators/ib/` | Exhibit |
| Audit / Advisory | `generators/audit/` | Annexe |
| Hedge Fund | `generators/hedge-fund/` | Exhibit |
| Equity Research | `generators/equity-research/` | Figure |
| Transverse | `generators/common/` | Annexe |
