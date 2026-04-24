# Plan de remédiation « fond en comble » — BDD-DOC

_Date : 24 avril 2026_

## Objectif

Traiter l’ensemble des risques identifiés sur BDD-DOC avec une approche **systémique**, en allant au-delà des correctifs ponctuels :
- sécurité opérationnelle,
- qualité de code VBA,
- fiabilité des flux métiers,
- robustesse à grande volumétrie,
- gouvernance et exploitabilité long terme.

---

## 0) Principes d’exécution (non négociables)

1. **Aucune régression fonctionnelle** : toute correction est accompagnée d’un test de non-régression.
2. **Traçabilité complète** : chaque changement = ticket + commit + preuve de test.
3. **Source de vérité unique** pour les exports VBA.
4. **Observabilité minimale** : erreurs et opérations critiques historisées.
5. **Sécurité pragmatique** : protection Excel ≠ sécurité forte ; secrets hors code.

---

## 1) Lot A — Sécurité et gouvernance technique

## A.1 Retrait des secrets en dur

### Constat
- `MDP_DEV` est présent en clair dans le code VBA.

### Actions
- Supprimer le secret en dur du dépôt.
- Mettre en place une valeur injectée via configuration locale (non versionnée).
- Rotation immédiate du mot de passe actuel.
- Ajouter une règle de revue bloquante : « aucun secret en clair ».

### Livrables
- Module de lecture de configuration sécurisé (best effort VBA).
- Gabarit `config.example` documenté (sans secret).
- Check automatique de détection de patterns sensibles (`password`, `token`, etc.).

### Critères d’acceptation
- Aucun secret détecté dans `src/` ni `vba_import/`.
- Démarrage classeur OK sans modification manuelle du code.

---

## A.2 Droits et mode développeur

### Constat
- Le mode dev manipule UI/protections, avec dépendances à l’environnement utilisateur.

### Actions
- Clarifier la politique d’accès (rôles : agent / dev / admin).
- Isoler les fonctions d’administration dans un module dédié.
- Ajouter des garde-fous : refus propre si prérequis non satisfaits.

### Livrables
- Matrice des rôles et permissions.
- Journal d’activation/désactivation du mode développeur.

### Critères d’acceptation
- Toute tentative d’accès non autorisée est tracée.
- Les transitions ON/OFF sont déterministes et testées.

---

## 2) Lot B — Architecture et maintenabilité VBA

## B.1 Refactor du module `Base`

### Constat
- Forte concentration de logique événementielle + métier dans un module volumineux.

### Actions
- Transformer `Base` en orchestrateur léger.
- Extraire en modules dédiés :
  - validation conformité,
  - styles/couleurs,
  - gestion undo/redo,
  - ouverture documentaire.
- Réduire la profondeur de conditions dans les événements.

### Livrables
- Nouveau découpage modulaire documenté.
- Cartographie des dépendances entre modules.

### Critères d’acceptation
- Baisse mesurable de complexité (taille procédures, responsabilités séparées).
- Scénarios métier inchangés côté utilisateur.

---

## B.2 Standard de gestion d’erreur

### Constat
- Usage hétérogène de `On Error`, dont des segments permissifs.

### Actions
- Standard unique :
  - `On Error GoTo Handler` en entrée,
  - bloc `SortiePropre`,
  - journalisation structurée (module, procédure, err.number, err.description, contexte).
- Limiter `Resume Next` aux micro-zones explicitement justifiées.

### Livrables
- Guide d’erreur VBA interne.
- Module `zDocLog` centralisé.

### Critères d’acceptation
- 100% des procédures critiques conformes au standard.
- Diminution des erreurs silencieuses.

---

## 3) Lot C — Données, synchronisation et intégrité

## C.1 Refonte GitSync (portable et fiable)

### Constat
- Couplage fort à des chemins locaux figés.

### Actions
- Externaliser la configuration de chemin repo.
- Ajouter un mécanisme de découverte guidée si absent.
- Valider l’encodage/fin de ligne de sortie.

### Livrables
- `GitSyncConfig` central.
- Script/check de validation pré-commit pour cohérence export.

### Critères d’acceptation
- Fonctionnement sur plusieurs postes sans modification du code.
- Export/import reproductible.

---

## C.2 Résolution définitive `src` vs `vba_import`

### Constat
- Double arborescence potentiellement divergente.

### Actions
- Définir une source canonique officielle.
- Mettre une normalisation automatique (BOM, CRLF/LF, fins de fichiers).
- Bloquer les commits si drift détecté.

### Livrables
- Politique de gestion des deux dossiers.
- Script de vérification de drift.

### Critères d’acceptation
- Diff résiduel expliqué et attendu.
- Zéro divergence non justifiée.

---

## C.3 Durcir l’intégrité de synchro métier

### Constat
- Synchronisation robuste mais encore dépendante d’hypothèses implicites (noms classeurs/onglets, contexte Excel).

### Actions
- Introduire des validations préflight exhaustives.
- Ajouter un rapport de synchro enrichi (durée, volumes, anomalies, rollback).
- Encadrer les cas d’échec partiel (reprise/retry contrôlé).

### Livrables
- Check-list préflight exécutable.
- Journal de synchro historisé.

### Critères d’acceptation
- Échec explicite et actionnable.
- Aucune corruption de données en cas d’interruption.

---

## 4) Lot D — Performance et stabilité à grande échelle

## D.1 Campagne de benchmarks reproductibles

### Actions
- Construire un protocole de test avec jeux de données représentatifs (petit / moyen / massif).
- Mesurer temps des workflows critiques : ouverture, édition conformité, filtres, synchronisation, sauvegarde.

### Livrables
- Tableau de bord baseline avant/après.
- Seuils de performance cibles.

### Critères d’acceptation
- Temps d’exécution stables (écart max défini).
- Aucune dégradation sur les scénarios principaux.

---

## D.2 Réduction des dépendances UI fragiles

### Actions
- Diminuer l’usage de `ActiveWindow`/`SendKeys`/états implicites.
- Mettre des fallback robustes pour sessions atypiques.

### Livrables
- Matrice de compatibilité contexte utilisateur (mono/multi écran, sessions distantes).

### Critères d’acceptation
- Comportement cohérent quel que soit le contexte.

---

## 5) Lot E — Qualité, tests et exploitation

## E.1 Stratégie de tests multi-niveaux

### Niveaux
1. **Smoke tests macro** (ouverture, saisie conformité, sauvegarde, synchro).
2. **Tests d’intégration fonctionnelle** (flux métier complet).
3. **Tests de non-régression ciblés** (bugs historiques).

### Livrables
- Catalogue de tests versionné.
- Modèle de compte-rendu de recette.

### Critères d’acceptation
- Tout correctif est lié à au moins un test.
- Historique de résultats conservé.

---

## E.2 Documentation d’exploitation

### Actions
- Rédiger un runbook support :
  - prérequis poste,
  - autorisations VBProject,
  - procédures de dépannage,
  - incidents fréquents et remédiation.

### Livrables
- `RUNBOOK_SUPPORT.md`
- `GUIDE_MAINTENANCE.md`

### Critères d’acceptation
- Un nouveau mainteneur peut opérer sans connaissance implicite.

---

## 6) Gouvernance du programme

## Organisation
- **Comité hebdomadaire** : suivi des risques / arbitrages.
- **Cycle en lots** : A → B → C → D → E (avec chevauchement contrôlé).
- **Definition of Done** stricte : code + test + doc + preuve.

## Indicateurs de pilotage
- Taux de procédures critiques alignées au standard d’erreur.
- Nombre d’erreurs silencieuses détectées.
- Taux de drift `src/vba_import`.
- Temps moyen des workflows critiques.
- Taux de succès synchro sans intervention manuelle.

---

## 7) Planning réaliste (fond en comble)

- **Phase 1 (2-3 semaines)** : Lot A + démarrage C.1
- **Phase 2 (3-5 semaines)** : Lot B complet + C.2
- **Phase 3 (2-4 semaines)** : C.3 + D.1 + D.2
- **Phase 4 (1-2 semaines)** : Lot E + stabilisation finale

> Durée totale estimée : **8 à 14 semaines** selon disponibilité et profondeur de recette.

---

## 8) Décision recommandée

Vu ta demande (« traiter de fond en comble »), la bonne approche est un **programme de remédiation complet** et non une série de patchs isolés.

La priorité absolue reste :
1. Sécuriser les secrets,
2. Rendre la chaîne GitSync portable/reproductible,
3. Mettre une base qualité (erreurs, tests, runbook),

puis industrialiser le refactor sans casser le métier.
