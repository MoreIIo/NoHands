# Extension : Excel ↔ Logiciel interne (automatisation)

Cette extension Chrome permet d'automatiser le transfert de valeurs entre un
fichier Excel (ou CSV) et une application web interne, sans droits admin ni
installation de logiciel externe.

## 1. Installation (aucun droit admin nécessaire)

1. Dézippez le dossier `extension/`.
2. Ouvrez Chrome et allez sur `chrome://extensions`.
3. Activez le bouton **"Mode développeur"** en haut à droite.
   *(Si ce bouton est grisé/absent, votre IT a désactivé cette fonctionnalité
   via une politique Chrome Enterprise — dans ce cas, demandez à l'IT
   d'autoriser le "mode développeur" ou l'installation de cette extension
   spécifique. Cela ne nécessite pas de droits admin sur le PC lui-même,
   uniquement une autorisation Chrome.)*
4. Cliquez sur **"Charger l'extension non empaquetée"** ("Load unpacked").
5. Sélectionnez le dossier `extension/`.
6. L'icône de l'extension apparaît dans la barre d'outils Chrome.

## 2. Ouverture du panneau

- Cliquez sur l'icône de l'extension : un panneau latéral s'ouvre à droite du
  navigateur et **reste ouvert** même quand vous cliquez sur la page (contrairement
  à un popup classique) — c'est indispensable pour choisir les champs à l'étape 4.

## 3. Utilisation, étape par étape

### Étape 1 — Charger le fichier Excel
Cliquez sur "Parcourir" et choisissez votre fichier `.xlsx`, `.xls` ou `.csv`.
S'il contient plusieurs feuilles, choisissez la bonne dans la liste déroulante.

### Étape 2 — Mapping des colonnes
Donnez un nom simple à chaque colonne utile. Exemple basé sur votre cas :

| Nom du champ      | Colonne |
|-------------------|---------|
| nom               | J       |
| prenom            | A       |
| valeur_recherche  | B       |

### Étape 3 — Conditions (filtrage des lignes)
Ajoutez une condition pour ignorer certaines lignes. Exemple :
`colonne E` `= égal à` `NON` → la ligne sera ignorée.
Vous pouvez ajouter plusieurs conditions (elles sont combinées en "OU" :
si une seule condition matche, la ligne est ignorée).

### Étape 4 — Configurer l'action sur le logiciel interne
1. **Ouvrez l'onglet** du logiciel interne dans Chrome (gardez le panneau latéral ouvert à côté).
2. Cliquez sur **"🎯 Choisir sur la page"** à côté de "Champ de recherche", puis
   cliquez directement sur le champ de saisie du logiciel interne où l'on tape
   la recherche. Le sélecteur CSS se remplit automatiquement.
3. Choisissez dans la liste déroulante quelle colonne (via son nom du mapping)
   doit être tapée dans ce champ.
4. Choisissez comment valider la recherche : "Appuyer sur Entrée" ou "Cliquer sur un bouton"
   (dans ce cas, utilisez aussi le bouton 🎯 pour sélectionner le bouton).
5. Réglez le délai d'attente après validation (le temps que le logiciel interne
   affiche le résultat, ajustez selon la lenteur du site — 1000 à 2000 ms est un bon départ).
6. Pour chaque information à récupérer sur la page de résultat, cliquez sur
   "+ Ajouter un résultat à récupérer", utilisez 🎯 pour cliquer sur l'élément
   contenant la valeur (ex: le numéro affiché), puis indiquez la colonne Excel
   où l'écrire.

### Étape 5 — Exécution
- Indiquez la ligne de départ (2 si la ligne 1 est l'en-tête) et éventuellement
  une ligne de fin.
- Cliquez sur **"▶ Démarrer"**. Le panneau va, pour chaque ligne non filtrée :
  taper la valeur de recherche, valider, attendre, lire les résultats, les écrire
  dans le tableau en mémoire.
- Le journal en bas affiche chaque ligne traitée (succès / ignorée / erreur).
- Vous pouvez cliquer sur "■ Arrêter" à tout moment.
- **"💾 Sauver config"** / **"📂 Charger config"** permettent de garder votre
  configuration (mapping, sélecteurs, conditions) d'une session à l'autre —
  utile si vous refaites la même tâche chaque jour avec un nouveau fichier Excel.

### Étape 6 — Télécharger le résultat
Cliquez sur **"⬇ Télécharger le fichier mis à jour"** : un fichier `.xlsx`
mis à jour est téléchargé dans votre dossier de téléchargements habituel.

## 4. Limites et points de vigilance

- **Site en `iframe`** : si le champ de recherche du logiciel interne est dans
  une `iframe`, le sélecteur CSS classique peut ne pas fonctionner directement.
  Dites-le-moi si c'est le cas, une adaptation est possible.
- **Anti-bot / CAPTCHA** : si le logiciel interne bloque les remplissages
  automatiques de champs, il faudra adapter la méthode de saisie.
- **Changements de mise en page** : si l'équipe qui gère le logiciel interne
  change l'interface, les sélecteurs CSS enregistrés peuvent casser — il suffira
  de refaire un clic avec 🎯 sur le nouvel élément.
- **Validez avec votre IT/manager** que l'automatisation de ce logiciel interne
  est autorisée par la politique de votre entreprise avant un usage en production.
- Cette extension ne transmet aucune donnée à un serveur externe : tout se passe
  localement dans votre navigateur (lecture/écriture du fichier Excel et lecture
  de la page web se font uniquement sur votre machine).

## 5. Pour aller plus loin

Si vous voulez que je pousse plus loin (par ex. gérer plusieurs valeurs à écrire
en une seule action, gérer les popups/nouveaux onglets ouverts par le logiciel
interne, ou ajouter une pause automatique en cas de CAPTCHA détecté), donnez-moi
plus de détails sur le comportement exact du logiciel interne et j'adapterai le code.
