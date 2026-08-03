# V10 — Mise en conformité "tout est piloté par l'Excel"

## Constat de départ

L'outil (V9) était déjà largement paramétré via `Parametrage_Outil_BIM_V1.xlsx`
(14 feuilles), mais **plusieurs textes métier étaient dupliqués et codés en
dur dans le Python**, avec une priorité qui écrasait silencieusement les
valeurs de l'Excel. Concrètement, modifier une cellule dans `10_Blocs` ou
`50_Questions` n'avait souvent **aucun effet visible** sur les livrables,
car le code utilisait un dictionnaire Python à la place.

Cinq foyers de duplication ont été identifiés et corrigés :

| Fichier | Contenu codé en dur | Écrasait |
|---|---|---|
| `bim_model.py` | `BLOCK_TITLES`, `BLOCK_DEMANDS`, `BLOCK_QUESTIONS`, `COMMERCIAL_PROPOSALS`, `STATUS_LABELS`, actions internes | `titre_bloc`, `lecture_entreprise`, `question_bim_manager` de `10_Blocs` |
| `content_rules.py` | `QUESTION_OVERRIDES` (7 questions B09/B06B/B12) | Le texte des mêmes questions dans `50_Questions` |
| `reporting_v2.py` | `PLAIN_TITLES` (3 occurrences, dont 2 dans du code mort) | Idem `titre_bloc` |
| `dashboard_template.py` | 4 anciennes définitions de `build_dashboard()` (~1600 lignes de code mort, jamais exécutées mais contenant un mapping de titres codé en dur en plus) | — |
| `outil_bim_v1.py` | Liste `_COLONNES_LIBRES` incomplète (substitution des `{variables}`) | — |

## Ce qui a changé

### 1. Excel `Parametrage_Outil_BIM_V1.xlsx`
- **`10_Blocs`** : 4 nouvelles colonnes ajoutées (X à AA), une par bloc (B01→B12) :
  - `proposition_commerciale` — texte à la 3ᵉ personne utilisé dans le
    paragraphe d'offre quand le statut est CONFIRMÉE.
  - `action_interne_standard` — action de sécurisation du Plan d'Actions.
  - `action_interne_documentaire` — variante si le lot est "documentaire"
    (vide = utilise `action_interne_standard`).
  - `action_interne_capacite` — action déclenchée quand l'auto-évaluation
    révèle un écart de capacité.
- **`16_Axes_Config`** : nouvel axe `statut_public` (CONFIRMED / CLARIFY /
  UNDEMONSTRATED / EXCLUDED) qui pilote désormais les libellés et l'ordre
  d'affichage des statuts, au lieu du dictionnaire `STATUS_LABELS` codé en dur.
- **`50_Questions`** : les 7 questions reformulées (B09_Q1-3, B06B_Q1-2,
  B12_Q1-2) qui n'existaient qu'en dur dans `content_rules.py` sont
  maintenant écrites directement dans la feuille — c'est ce texte-là qui
  s'affichera, et vous pouvez l'éditer.

### 2. Code Python
- Tous les dictionnaires codés en dur ont été **renommés en `_FALLBACK_*`**
  et rétrogradés en filet de sécurité : ils ne sont utilisés que si une
  cellule Excel est vide, jamais en priorité sur l'Excel.
- `ensure_public_model()` accepte maintenant un paramètre `axes_config` pour
  lire les libellés de statut depuis `16_Axes_Config`.
- Les deux listes `_COLONNES_LIBRES` (une par chemin de génération, web et
  PDF) ont été complétées avec les 4 nouvelles colonnes, pour que les
  `{variables}` détectées (`{lot}`, `{plateforme}`, etc.) soient aussi
  substituées dans ces nouveaux textes.
- **Nettoyage** : suppression de ~1600 lignes de code mort dans
  `dashboard_template.py` (4 anciennes versions de `build_dashboard()`
  jamais appelées, qui portaient elles-mêmes un mapping de titres codé en
  dur — source de confusion pour toute personne modifiant le fichier).

## Vérifications effectuées

- ✅ Cohérence des identifiants de blocs entre les feuilles `10_Blocs`,
  `50_Questions`, `60_Reponses` : les 13 blocs (B01→B12, avec B06A/B06B)
  sont identiques partout — aucun problème trouvé à ce niveau.
- ✅ Pipeline complet réexécuté avec les PDF d'exemple du projet
  (`samples/Convention_test.pdf` + `CCTP_test.pdf`) avant/après modification.
- ✅ Répartition des statuts (Confirmée / À confirmer / Non démontrée / Non
  applicable) strictement identique avant/après — la logique de décision
  n'a pas été touchée, seul le texte affiché change.
- ✅ Les nouveaux titres de blocs (issus de `10_Blocs`, plus longs et plus
  précis que les anciens) s'affichent et se mettent en page correctement
  dans les PDF (aucune page supplémentaire, aucun texte tronqué).
- ✅ Les nouvelles colonnes `proposition_commerciale` /
  `action_interne_standard` / `action_interne_capacite` apparaissent
  correctement dans `Plan_Actions_Offre_BIM.pdf/html`.
- ✅ Tous les fichiers Python compilent sans erreur.

## Ce que ça change concrètement pour vous

- Les titres de blocs affichés dans `Rapport_BIM_Complet` sont maintenant
  ceux de la colonne `titre_bloc` de `10_Blocs` (plus longs et détaillés
  que les anciens titres courts codés en dur) — vérifiez qu'ils vous
  conviennent, vous pouvez les raccourcir directement dans l'Excel.
- Vous pouvez maintenant éditer `10_Blocs` (textes, actions, propositions
  commerciales) ou `50_Questions` et voir l'effet directement dans les
  livrables régénérés, sans toucher au code.
- Le dossier `exemples_regeneres/` contient les 5 livrables régénérés avec
  ce code + cet Excel, à partir des PDF d'exemple du projet — comparez-les
  à vos anciens exports pour voir la différence.

## Limites connues / pistes non traitées

- Les couleurs et polices des PDF (`STYLE_PROFILES` dans `reporting_v2.py`,
  `STATUT_META` dans `outil_bim_v1.py`) restent codées en dur : ce sont des
  choix de design (thème visuel), pas du contenu métier BIM, donc je les ai
  laissés en l'état. Dites-moi si vous voulez aussi les paramétrer.
- Je n'ai pas eu vos vrais documents (Convention + CCTP réels) pour
  régénérer exactement `Plan_Actions_Offre_BIM_11_.html` /
  `__2_.html` que vous aviez fournis — relancez l'outil avec vos documents
  pour obtenir la version corrigée de vos livrables réels.
