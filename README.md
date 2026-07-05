# NoHands OSA — Extension Chrome

Fusion des projets **NoHands** (remplissage de formulaires depuis Excel) et
**OSA** (récupération de résultats web vers Excel) en une seule extension à
deux modes :

- **Saisie** : choisis une ligne de ton fichier Excel, l'extension remplit les
  formulaires du site (attributs `name`), y compris les popups du même site,
  les selects, checkboxes, dates, et les IBAN découpés en plusieurs champs.
- **Extraction** : pour chaque ligne du fichier, l'extension tape une valeur
  dans le champ de recherche du site, valide, attend, lit les résultats
  (sélecteur CSS ou ligne de tableau) et les écrit dans le fichier, que tu
  télécharges mis à jour à la fin.

## Installation (aucun droit admin nécessaire)

1. Ouvre Chrome → `chrome://extensions`.
2. Active le **Mode développeur** (en haut à droite).
3. **Charger l'extension non empaquetée** → sélectionne ce dossier.
4. Clique sur l'icône : le panneau latéral s'ouvre (raccourci `Ctrl+Shift+N`).

> Le dossier `_backup_avant_fusion/` contient l'ancienne version OSA telle
> qu'elle était avant la fusion. Tu peux le supprimer quand tout te convient.

## 1. Onglet Données

- **Sources** : fichier `.xlsx`/`.xls`/`.csv`, tableau collé depuis Excel
  (Ctrl+V), ou JSON (collé ou fichier).
- **En-têtes** : détection automatique de la 1re ligne (forçable en Oui/Non).
- **Colonnes** :
  - avec en-têtes → les noms de colonnes sont lus dans le fichier ;
  - sans en-têtes → choisis un **modèle de colonnes** (PROP, LOTS et BAIL sont
    fournis par défaut, hérités de NoHands) ou reste en lettres A, B, C…
  - Si le nombre de colonnes collées correspond exactement à un seul modèle,
    il est appliqué automatiquement.
- **Modèles** : bouton « Gérer » → créer, modifier, supprimer des modèles
  (un nom de colonne par ligne), ou en créer un depuis les colonnes actuelles.

## 2. Onglet Saisie (ex-NoHands)

1. Sélectionne la ligne à saisir (filtre + navigation ◀ ▶).
2. Les champs de la ligne s'affichent avec un bouton copier (IBAN formaté
   automatiquement, trimestres normalisés en T1/T2…).
3. **Configurer le mapping** : pour chaque colonne, indique le(s) `name` des
   inputs du site. Deux façons de les récupérer :
   - clic droit sur le champ du site → **Copier le nom de l'input** ;
   - bouton 🎯 dans le mapping → clique directement sur le champ.
   Le mapping est mémorisé par modèle / jeu de colonnes.
4. **Champs personnalisés** : paires name→valeur fixes (import/export JSON).
5. **Remplir le formulaire** : remplit tous les onglets ouverts du même site
   (popups compris). Les contenus chargés dynamiquement (ASP.NET UpdatePanel)
   sont re-remplis automatiquement pendant 30 s.

## 3. Onglet Extraction (ex-OSA)

1. **Conditions** : ignorer certaines lignes (ex : colonne X = "NON").
2. **Champs de recherche** : sélecteur CSS (bouton 🎯) + colonne à taper.
3. **Validation** : touche Entrée ou clic sur un bouton, puis délai d'attente.
4. **Résultats à récupérer** : sélecteur CSS, ou « ligne de tableau » (compare
   une colonne du fichier à une cellule de chaque ligne du tableau de la page,
   puis extrait une autre cellule). Destination : colonne existante, lettre,
   ou **nouvelle colonne** créée à droite du tableau.
5. **Exécution** : plage de lignes, ▶ Démarrer / ■ Arrêter, journal et
   récapitulatif détaillé.

## 4. Onglet Export

Télécharge le fichier mis à jour en `.xlsx` (même feuille remplacée si le
fichier vient d'Excel) ou en `.json`.

## Profils

La barre du haut permet d'enregistrer des **profils** nommés : mapping,
champs personnalisés, configuration d'extraction, mode en-têtes et modèle.
Utile si tu alternes entre plusieurs tâches récurrentes. L'ancienne
configuration OSA sauvegardée est migrée automatiquement dans le profil
« OSA (importé) ».

## Test local

Ouvre `test.html` dans Chrome : la page contient un formulaire complet pour
tester la Saisie (avec IBAN découpé) et une recherche simulée avec tableau de
résultats pour tester l'Extraction.

## Limites et vigilance

- **Iframes** : si le champ cible est dans une iframe, le sélecteur CSS simple
  peut ne pas suffire (le remplissage par `name` fonctionne dans les frames).
- **Anti-bot / CAPTCHA** : non contourné, par conception.
- **Changements d'interface du site** : re-sélectionne les éléments avec 🎯.
- Aucune donnée n'est envoyée à un serveur : tout reste dans ton navigateur.
- Vérifie que l'automatisation du logiciel interne est autorisée par ta
  politique d'entreprise avant un usage en production.

## Structure du projet

```
├── manifest.json      Manifest V3 (side panel, content script, menus)
├── background.js      Service worker : panneau, menu contextuel, injection
├── content.js         Moteur de remplissage (mode Saisie)
├── sidepanel.html     Interface (Données / Saisie / Extraction / Export)
├── sidepanel.css      Styles
├── sidepanel.js       Logique complète
├── lib/xlsx.full.min.js  SheetJS (lecture/écriture Excel)
├── icons/             Icônes de l'extension
└── test.html          Page de test locale des deux modes
```

**Version** : 2.0.0 — fusion NoHands + OSA.
