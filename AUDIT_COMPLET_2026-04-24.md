# Audit complet de BDD-DOC

_Date de l'audit : 24 avril 2026_

## 1) Périmètre

Audit statique du dépôt `BDD-DOC` (modules VBA exportés) sur :
- architecture et maintenabilité,
- robustesse (gestion erreurs / états Excel),
- sécurité (protection, mots de passe, accès fichiers/système),
- performance (volumétrie, patterns événementiels),
- exploitabilité Dev (synchronisation Git / reproductibilité).

## 2) Méthodologie

- Inventaire des modules `src/` et `vba_import/`.
- Revue ciblée des modules critiques :
  - `Base.txt` (événements feuille),
  - `ThisWorkbook.txt` (cycle de vie classeur),
  - `zDocImportBDD.bas` (synchronisation métier),
  - `zDocGitSync.bas` (export/import Git),
  - `zDocDeveloppeurOnOFF.bas` (protection et gouvernance),
  - `zDocOuvertureWord.bas`, `zDocOuvertureExcel.bas`, `zDocWindows.bas`.
- Contrôles semi-automatisés : présence `Option Explicit`, patterns `On Error`, API sensibles (`Kill`, `CreateObject`, `Application.Run`, etc.), divergence `src` vs `vba_import`.

## 3) Forces observées

1. **Structuration fonctionnelle claire** : séparation par modules métier (import, UI, sécurité, ouverture documents, audit perf).
2. **Usage quasi systématique de `Option Explicit`** (discipline de base VBA respectée).
3. **Prise en compte de la volumétrie** sur `Base` (fast-path, contrôles de zones, usage matriciel par `Area`).
4. **Restauration d'états Excel** souvent bien pensée (`ScreenUpdating`, `EnableEvents`, `Calculation`).
5. **Existence d'un module `AuditPerformance`** utile pour diagnostiquer le classeur côté production.

## 4) Risques et constats prioritaires

## Critique (à corriger en priorité)

### C1 — Secret en clair dans le code
- Le mot de passe développeur est stocké en dur (`MDP_DEV = "11223344"`).
- Impact : fuite facile via export VBA, contournement de la protection des feuilles/classeur.
- Recommandation :
  - sortir ce secret du code (au minimum valeur injectée hors dépôt),
  - mettre en place une rotation immédiate du mot de passe,
  - considérer la protection comme une **barrière UX**, pas une sécurité forte.

### C2 — Couplage fort à un poste local pour GitSync
- Les chemins repo sont codés en dur pour deux machines (`C:\Users\...` et `D:\...`).
- Impact : non-portabilité, échecs silencieux sur tout autre environnement.
- Recommandation :
  - externaliser via cellule de config, variable d'environnement, ou dialogue de sélection dossier.

## Élevé

### E1 — Double source de vérité (`src` et `vba_import`) systématiquement divergentes
- Le dépôt maintient deux arborescences quasi miroir mais les fichiers diffèrent globalement (BOM/fin de ligne et possiblement contenu).
- Impact : risque de dérive, revue Git bruitée, confusion sur la “bonne” version.
- Recommandation :
  - définir explicitement la source canonique,
  - normaliser l'encodage/fin de ligne,
  - automatiser la vérification de synchro avant commit.

### E2 — Gestion d'erreur hétérogène et partiellement permissive
- Usage fréquent de `On Error Resume Next` dans les flux sensibles (IO, protection, objets COM).
- Impact : masquage d'anomalies, diagnostics difficiles en production.
- Recommandation :
  - limiter `Resume Next` à des blocs très courts et immédiatement vérifiés,
  - journaliser centralement les erreurs fonctionnelles.

### E3 — Surface d'exécution dynamique importante
- Appels dynamiques `Application.Run`, automation `CreateObject`, manipulations registry/FS.
- Impact : comportement difficile à tracer, dépendance forte à l'environnement Windows/Office local.
- Recommandation :
  - encapsuler les points d'entrée dynamiques,
  - centraliser les contrôles prérequis et la journalisation.

## Moyen

### M1 — Module `Base` très chargé (événementiel + logique métier)
- `Base.txt` concentre beaucoup de responsabilités (UI, validation, undo, couleurs, double-clic d'ouverture doc).
- Impact : complexité cognitive, risque de régression lors d'évolutions.
- Recommandation :
  - extraire la logique métier vers modules dédiés, garder les événements comme orchestrateurs.

### M2 — Conversion URL/chemin potentiellement fragile sur UTF-8
- `UrlDecodeUtf8` reconstruit une chaîne via `StrConv(..., vbUnicode)` sur octets ; cas accentués/non ASCII à valider fortement.
- Impact : ouverture de documents pouvant échouer sur chemins internationaux.
- Recommandation :
  - ajouter tests de non-régression sur chemins accentués/UNC/espaces encodés.

### M3 — Dépendance à l'UI / `ActiveWindow` / `SendKeys`
- Certaines fonctionnalités de mode développeur ou mise en forme reposent sur l'état de fenêtre active.
- Impact : instabilité possible selon contexte utilisateur (multi-fenêtres, verrouillage poste, sessions distantes).
- Recommandation :
  - réduire les dépendances implicites à l'UI active,
  - protéger les opérations non déterministes.

## 5) Plan d'action recommandé

## Sprint 1 (sécurisation / fiabilisation)
1. Retirer le secret en clair (`MDP_DEV`) + rotation du mot de passe.
2. Externaliser la configuration des chemins `zDocGitSync`.
3. Ajouter un log technique central (feuille log ou fichier local) pour erreurs critiques.

## Sprint 2 (qualité code)
4. Refactor `Base.txt` : extraction en services (validation conformité, style, ouverture documents).
5. Standardiser la politique d'erreur (`GoTo Handler` + sorties propres).
6. Normaliser `src`/`vba_import` (encodage, source canonique, check automatique).

## Sprint 3 (industrialisation)
7. Ajouter un pipeline de contrôles statiques (lint simple scripts + checks de drift).
8. Ajouter un mini jeu de tests fonctionnels macro (smoke tests guidés).
9. Documenter le runbook support (prérequis Office, droits VBProject, chemins, dépannage).

## 6) Synthèse exécutive

Le projet est **fonctionnel, structuré et orienté usage terrain**, avec de bonnes bases VBA (clarté des modules, gestion d'états Excel, prise en compte de la volumétrie).

Les principaux risques sont surtout de **gouvernance technique** et **sécurité opérationnelle** : secret en dur, dépendances poste local, divergence structurelle `src`/`vba_import`, et gestion d'erreurs trop permissive par endroits.

Avec 2 à 3 itérations ciblées, la base peut atteindre un niveau nettement plus robuste et maintenable sans remise à plat complète.
