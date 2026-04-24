# Outils qualité dépôt

## Vérification de synchronisation `src/` vs `vba_import/`

```bash
python tools/check_src_sync.py
```

- Mode par défaut = **report non bloquant** (retour `0`) pour visualiser les drifts.
- Pour échouer si drift détecté :

```bash
python tools/check_src_sync.py --enforce
```

- Option stricte :

```bash
python tools/check_src_sync.py --enforce --strict
```

En mode strict, même les diffs de normalisation BOM/EOL provoquent un échec.

## Correction automatique (safe)

Pour corriger automatiquement les diffs de normalisation **uniquement** (BOM/EOL) :

```bash
python tools/check_src_sync.py --fix-normalization
```

Ce mode ne corrige pas les drifts de contenu réel ni les binaires différents.

## Quand demander une extraction

Demander une nouvelle extraction est recommandé si :
- des modifications ont été faites directement dans le classeur Excel (éditeur VBA / UserForms),
- des fichiers `.frx` changent (binaire non réconciliable manuellement),
- le check montre des drifts de contenu réel sur plusieurs modules.

Pas nécessaire immédiatement si :
- seules des normalisations BOM/EOL sont en jeu (corrigeables via `--fix-normalization`).
