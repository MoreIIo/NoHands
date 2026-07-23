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
6. **Champs à autocomplétion asynchrone** (ex. code postal → ville, comptes
   bancaires) : détectés automatiquement (handler `searchResult`/`SearchStart`
   ou conteneur `search:…`). L'extension tape la valeur, attend les
   suggestions AJAX (polling 200 ms, timeout 8 s — jamais de délai fixe),
   choisit la bonne option en la comparant aux **autres colonnes de la ligne**
   (ex. la colonne Ville départage « 70600 - ARGILLIERES » de
   « 70600 - AUTREY... »), puis la clique pour déclencher `setDataFieldValue`
   (la ville liée se remplit alors côté site). Si la détection échoue sur un
   formulaire, préfixe le nom d'input par `ac:` dans le mapping ou les champs
   personnalisés (ex. `ac:body_x_city_x_Migrated_txtCP_x_CP`). L'option
   choisie apparaît dans le retour de remplissage ; si aucune suggestion
   n'apparaît, la valeur tapée reste et un avertissement est signalé.
7. **Formulaires « en cascade » (WebForms / `__doPostBack`)** : les menus dont
   le changement recharge la page (ex. Pays) sont détectés : l'extension
   attend la fin du rechargement partiel avant de passer au champ suivant.
   Les listes peuplées par le serveur (ex. Numéro de voie, vide au départ)
   sont attendues jusqu'à 8 s. Pour les champs texte qui doivent déclencher
   un rechargement à la sortie (ex. code postal → `initCacheresultCall`,
   voie), préfixe l'input par `pb:` dans le mapping
   (ex. `pb:body_x_txtVoie_x_Voie`) : blur ciblé puis attente du
   rechargement. Un champ déjà à la bonne valeur n'est jamais re-rempli
   (évite les boucles de postback). L'ordre de remplissage suit l'ordre des
   colonnes : place Pays avant Code postal, Code postal avant Voie, etc.
   Les marqueurs se cumulent (`pb:ac:monChamp`).

8. **Étape de scénario « SIGEO : saisir une adresse »** : automatise le
   formulaire d'adresse SIGEO (`popup.aspx/…/address_manage/{id}`) de bout en
   bout — navigation (ViewState frais), pays, CP (autocomplétion
   « searchResult » : la suggestion CP+ville est cliquée pour résoudre le
   **code commune** `Migrated_code`), voie, **n° de voie** (match strict du
   libellé dans la liste), champs optionnels, contrôles, puis « Enreg. et
   fermer ». Fallbacks : sélecteur de commune (popup interne ou fenêtre
   séparée), reprise idempotente après postback. **Simulation (dryRun) activée
   par défaut** : tout est rempli et contrôlé sans enregistrer — décocher pour
   écrire réellement. Résultat par ligne (OK / SIMULATION OK / ERREUR + ids
   résolus) écrit en option dans une colonne du fichier. Codes d'erreur :
   `AUTH_REQUIRED` (session expirée), `UNRESOLVED_CITY` (CP/ville incohérents),
   `NUM_VOIE_INTROUVABLE` (libellé absent), `VALIDATION` (refus serveur).

9. **Étape de scénario « Éditer par lots (cocher N → cliquer) »** : pour les
   pages qui limitent le nombre d'éléments traitables en une fois. L'étape
   coche les cases par paquets de N, clique un bouton entre chaque paquet,
   et recommence jusqu'à épuisement. Cas d'usage : l'édition de feuille de
   présence SIGEO (`syn_man_edition_feuille_presence`) — 21 clés de
   répartition, lots de 10 → le bouton « Editer » est cliqué 3 fois, donc
   3 `feuille.doc` téléchargées.

   Réglages : **tableau** (id ou sélecteur CSS, bouton 🎯 pour le désigner ;
   vide = toute la page), **bouton à cliquer**, **taille de lot**,
   **attente après le clic** (doit couvrir la génération + le
   téléchargement du fichier), **filtre** optionnel sur le libellé (texte,
   ou `/regex/i` — pratique pour écarter une clé dont le total est nul), et
   un garde-fou **lots max**. Le journal détaille chaque lot
   (`lot 2 : éléments 11-20 / 21 → clic`), et le bouton « Arrêter » du
   scénario interrompt la boucle entre deux lots.

   Points techniques : avant chaque clic, l'étape décoche tout puis coche
   uniquement la tranche voulue — aucun risque de cumul d'un lot sur
   l'autre. Le découpage suit l'ordre du document, donc il reste stable même
   si la page se recharge partiellement entre deux lots. Sur ces tables
   chaque case `id="X"` a un hidden miroir `id="hdnX"` (c'est *lui* qui est
   posté au serveur) : l'extension passe par le clic natif plutôt que par
   `checked = true`, resynchronise le miroir en filet de sécurité, et ne
   reclique jamais une case déjà dans le bon état.

10. **Exécution du scénario en parallèle sur plusieurs onglets** : le bouton
   « Démarrer sur plusieurs onglets » (section *Exécution du scénario*)
   répartit les lignes de la plage entre N onglets — N étant le nombre
   réglé dans *Saisie multi-onglets*, juste au-dessus. Chaque onglet pioche
   la ligne suivante dans une file commune et déroule le scénario complet
   dessus, indépendamment des autres : aucune ligne n'est traitée deux
   fois, et un onglet lent ne bloque pas les autres. Le journal préfixe
   chaque ligne par l'onglet (`O2 · L7 · Étape 1 — …`) et « Arrêter »
   interrompt tous les onglets.

   Les onglets déjà ouverts sur le site sont réutilisés ; les manquants
   sont créés en dupliquant l'onglet du formulaire, et l'exécution attend
   qu'ils aient fini de charger.

   Différence importante avec le mode simple : le **remplissage est
   cloisonné** à l'onglet qui traite la ligne (et aux popups qu'il a
   ouvertes). En mode simple, « Remplir » diffuse à tous les onglets du
   site, ce qui serait destructeur ici — chaque onglet écraserait la ligne
   des autres. La ligne active de l'interface ne suit pas la boucle, car
   plusieurs lignes avancent en même temps.

   À savoir : ce mode ne convient pas aux scénarios dont les étapes
   partagent un état global, en particulier les étapes **PDF β** qui
   travaillent sur le document actif de la Toolbox.

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

## 5. Onglet Toolbox (bêta)

Analyse locale de **PDF** (pdf.js embarqué, rien n'est envoyé sur internet) :

1. **Documents** : charge un ou plusieurs PDF ; le texte est extrait page par
   page (avec reconstruction des espaces pour les PDF mal encodés).
2. **Champs détectés** : motifs intégrés (n° ADEME/DPE, SIRET, SIREN, TVA,
   IBAN, dates, emails, téléphones, CP+ville, surfaces) + libellés
   « Libellé : valeur » (propriétaire, adresse, diagnostiqueur…). Clic sur une
   valeur = copie. Les champs cochés peuvent être écrits d'un coup dans la
   **ligne active** (colonnes du même nom, créées si besoin).
3. **Motifs personnalisés** : ajoute tes propres champs (libellé ou regex),
   mémorisés dans le navigateur.
4. **Scénario** : deux nouveaux types d'étape dans le scénario de saisie —
   **PDF β : vérifier un champ** (compare une donnée du PDF à une valeur fixe
   ou `{Colonne}` ; en cas d'écart : erreur, avertissement, saut d'étapes ou
   arrêt) et **PDF β : écrire un champ** (reporte la valeur dans une colonne).
   En boucle, le bon document est retrouvé par son **nom de fichier**
   (ex. `{N° DPE}`).

**OCR (bêta)** : bouton « Lancer l'OCR » sous la liste des documents — reconnaît
le texte des **pages scannées et des images** (Tesseract embarqué, 100 % local,
~5-20 s par page). Deux modes : *pages sans texte (auto)* pour les PDF scannés,
ou *toutes les pages* pour lire par ex. l'**étiquette énergie** d'un DPE (conso
`kWh/m²/an`, émissions `kgCO₂/m²/an`), qui est une image même dans les PDF
texte. Les champs sont re-détectés après l'OCR (motifs tolérants aux erreurs de
reconnaissance) et le texte OCR s'ajoute au texte extrait (sections
« — OCR page N — »).

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

## Mises à jour (depuis GitHub)

L'extension peut vérifier si une version plus récente est publiée sur le dépôt
GitHub `MoreIIo/NoHands` et afficher une bannière avec les notes de version.

1. Clique sur l'icône ⚙️ (barre des profils) → **Mises à jour**.
2. Colle un **token GitHub** en lecture seule, puis **Enregistrer le token**.
   - Le dépôt étant privé, un token est nécessaire pour lire la version publiée.
   - Recommandé : *fine-grained token* limité au dépôt `MoreIIo/NoHands`,
     permission **Contents : Read-only**.
   - Le token est stocké uniquement dans ce navigateur (`chrome.storage.local`),
     jamais dans le code ni poussé sur GitHub, et n'est envoyé qu'à
     `api.github.com`.
3. **Vérifier maintenant** compare la version installée à celle de
   `manifest.json` sur la branche `main`. Si une version supérieure existe, une
   bannière « Nouvelle version disponible » s'affiche avec un lien vers GitHub.

La vérification se relance automatiquement à l'ouverture du panneau (au plus une
fois toutes les 3 h ; désactivable via la case à cocher). Les notes proviennent
de la dernière *release* si elle existe, sinon des derniers messages de commit.

> L'extension étant chargée « non empaquetée », la mise à jour reste **manuelle** :
> récupère la nouvelle version sur GitHub puis recharge l'extension dans
> `chrome://extensions`. La bannière sert d'alerte, pas d'installateur automatique.

## Structure du projet

```
├── manifest.json      Manifest V3 (side panel, content script, menus)
├── background.js      Service worker : panneau, menu contextuel, injection
├── content.js         Moteur de remplissage (mode Saisie)
├── sidepanel.html     Interface (Données / Saisie / Extraction / Export / Toolbox)
├── sidepanel.css      Styles
├── sidepanel.js       Logique complète
├── update-check.js    Vérification de mise à jour via GitHub
├── lib/xlsx.full.min.js  SheetJS (lecture/écriture Excel)
├── lib/pdf.min.js     pdf.js (Mozilla) — analyse PDF de la Toolbox
├── lib/pdf.worker.min.js  Worker pdf.js
├── lib/ocr/           Tesseract.js (OCR local) : lib, worker, cœur wasm, langue fra
├── icons/             Icônes de l'extension
└── test.html          Page de test locale des deux modes
```

**Version** : 2.5.0 — étape de scénario **SIGEO : saisir une adresse**
(formulaire address_manage : résolution code commune via autocomplétion ou
sélecteur, n° de voie par libellé, dryRun, rapport par ligne).
2.2.0 — Toolbox : **OCR local** (Tesseract) pour PDF scannés et
images (étiquette énergie DPE : conso kWh/m²/an, émissions kgCO₂/m²/an).
2.1.0 — onglet **Toolbox (bêta)** : analyse de PDF (n° ADEME, SIRET,
noms, dates…), vérification/report des données dans le scénario de saisie.
