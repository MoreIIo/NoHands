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

> Le dossier `backup_avant_fusion/` (local, non publié) contient l'ancienne
> version OSA telle qu'elle était avant la fusion. Tu peux le supprimer quand
> tout te convient.

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
   Ils suivent **exactement la même logique que le mapping** : une variable
   `{Colonne}` reçoit la même valeur que si la colonne était mappée
   (formatage IBAN / trimestre, règles de valeurs), et le remplissage passe
   par le même moteur (marqueurs `ac:` / `pb:` / `nopb:` / `mg:`,
   autocomplétion, postbacks, erreurs). Idem pour l'étape « Saisir un champ ».
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
   Depuis la v2.8.2, les champs texte ayant un `onblur` inline (sans
   `__doPostBack`) reçoivent automatiquement la sortie de champ après
   remplissage, comme si tu cliquais ailleurs. `pb:` reste utile pour les
   champs dont le handler n'est pas visible dans le HTML ; `nopb:` désactive
   la sortie automatique sur un champ.

   **Champ SIGEO « Mandat MG »** (saisie d'opérations diverses) : ce n'est
   pas un input simple (sélecteur + valeur cachée soumise + postback qui
   recharge le Journal). Mappe la colonne « Mandat MG » sur `mg:` (ou
   `mg:body_x_selMan_x` ; choisir le champ avec 🎯 est aussi détecté
   automatiquement). Valeurs acceptées : `MG0396166`, `mg396166`, `396166`.
   Le mandat est **toujours saisi en premier**, quelle que soit la position
   de la colonne ; l'extension attend la fin du rechargement (8 s max) puis
   vérifie la valeur soumise et le libellé. En cas d'échec (entrée mal
   formée, mandat inexistant, champ absent, délai dépassé), **le reste du
   remplissage est abandonné** et l'étape de scénario passe en erreur.
   Aucun bouton (« Saisir les contreparties »…) n'est jamais cliqué.
   Fonctionne aussi dans l'étape « Saisir un champ » (champ `mg:`).

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
   **mode d'attente après le clic**, **filtre** optionnel sur le libellé
   (texte, ou `/regex/i` — pratique pour écarter une clé dont le total est
   nul), et un garde-fou **lots max**.

   Trois modes d'attente entre les lots : **délai fixe** (prévoir large :
   génération + téléchargement), **début du téléchargement** (enchaîne dès
   que le navigateur commence à recevoir le fichier — le plus rapide), ou
   **fin du téléchargement** (attend que le fichier soit complet — le plus
   sûr si le serveur est lent). Les deux modes basés sur le téléchargement
   utilisent l'accès aux téléchargements de Chrome (permission
   `downloads` du manifest). Si aucun fichier n'arrive dans le délai d'abandon
   (30 s par défaut), le lot est signalé « sans téléchargement » et la
   boucle continue. En mode multi-onglets, l'attente du téléchargement
   n'est pas fiable (Chrome ne dit pas quel onglet a déclenché le
   fichier) : un avertissement est journalisé. Le journal détaille chaque lot
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

10 bis. **Récap de fin de boucle et reprise** (v2.21.0) : à la fin de
   « Démarrer le scénario » ou « Démarrer sur plusieurs onglets », un
   encadré **Récap de la boucle** s'affiche sous le journal : nombre de
   lignes réussies / en erreur / vides / non traitées, temps passé,
   erreurs regroupées par étape, et tableau des **lots en erreur** (ligne,
   lot — repris des colonnes *lot / immeuble / mandat / réf / code /
   adresse*, sinon les deux premières cellules —, étape en échec, raison).
   Le détail de toutes les lignes est dépliable et **Exporter CSV** sort le
   tout pour Excel.

   L'avancement est enregistré après **chaque ligne**. Si la boucle s'arrête
   avant la fin (étape réglée sur « si erreur : stopper », bouton Arrêter,
   panneau fermé, extension rechargée), le récap le signale et propose
   **Reprendre à la ligne X** : seules les lignes non traitées sont
   rejouées, plus la ligne qui a provoqué l'arrêt — les lignes réussies ne
   sont pas refaites. **Relancer les erreurs** rejoue uniquement les lignes
   en erreur. Le récap se met à jour au fil des reprises et reste affiché
   (même après fermeture du panneau) jusqu'au ✕ ou au prochain démarrage.
   La reprise exige le même fichier (nom + nombre de lignes) que la boucle
   d'origine.

11. **Étape de scénario « Sélecteur SIGEO »** (bouton **+ Sélecteur SIGEO**) :
   dédiée aux champs SIGEO à autocomplétion (`SelectorControl` +
   `AutoCompletion`), où écrire la valeur ne suffit pas.

   | Réglage | Rôle |
   |---|---|
   | champ | `Mandat MG`, `Compte` (ligne de contrepartie N) ou `Personnalisé` |
   | ligne | n° de ligne de contrepartie (Compte uniquement) |
   | valeur | fixe ou `{Colonne}` (même traitement que le mapping) |
   | clé __ivCtrl | Personnalisé : clé du contrôle. Compte : modèle optionnel (`{n}`, `{n0}`, `{ctl}`), sinon détection auto |
   | mode | `direct` (id + `ctrl.add`, imposé pour Mandat MG) ou `recherche` (frappe + suggestion, imposé pour Compte) |
   | timeout | 8000 ms (Mandat MG) / 5000 ms (Compte) |
   | si erreur | **stopper** tout le scénario, ou ignorer la ligne du batch et continuer |

   Validation : une seule valeur sélectionnée (`SelectedValues.length === 1`)
   et texte affiché qui commence par la valeur ; suggestions ambiguës refusées
   avec la liste des candidats. **Tester l'étape** l'exécute seule sur
   l'onglet actif. Chaque exécution est inscrite dans le **journal des
   sélecteurs SIGEO** (ligne, valeur, id rattaché, libellé, statut ; export
   CSV). Garde-fous : une étape Compte exige un Mandat MG réussi plus tôt
   dans le scénario ; avertissement si une étape Débit / Crédit / TVA d'une
   ligne précède le Compte de cette ligne. Aucun clic d'enregistrement.
   Le bouton **Modèle OD** ajoute le scénario « OD – Saisie contrepartie »
   (Mandat MG → clic « Saisir les contreparties » → Compte ligne 1 → Débit →
   Crédit → Libellé), à compléter avec 🎯.

12. **Étape de scénario « SIGEO : saisir un DPE »** (bouton **Modèle DPE**) :
   remplit le formulaire **Diagnostic** de SIGEO (popup ouverte en onglet,
   `…/popup.aspx/fr/sig_imm/diag_manage/{id}/10/{idDiag}`, `0` en création),
   clique **Enregistrer** puis joint le **PDF du DPE** (description « DPE »).

   **Colonnes** (modèles `{Colonne}` modifiables, bloc « Colonnes du
   fichier ») : `type_local` (habitation → 11 par défaut, hors habitation →
   15), `code_fournisseur`*, `date_etabli`*, `date_validite` (jamais saisie :
   comparée au calcul SIGEO, simple avertissement si elle diffère),
   `date_debut_validite`, `numero_ademe`, `conso_ep`*, `emission_ges`*,
   `lettre_ep_forcee` / `lettre_ges_forcee` (lettres **du DPE** : voir
   ci-dessous), `altitude_800m` / `nouvelles_valeurs` (oui/non), `cout_min`,
   `cout_max`, `date_prix`, `resultat`, `commentaire`, `fichier_dpe`, `lot`
   (affichage).
   Le bouton « Colonnes de l'Analyse DPE » reprend celles écrites par
   l'extraction *Analyse DPE* (`{N° FU}`, `{Établi le}`, `{Classe énergie}`,
   `{DPE fichier}`…) ; « Fichier DPE extraction » celles du fichier prêt de
   l'extraction *DPE extraction* (+ URL `{URL fiche DPE}`). « Modèle DPE »
   choisit tout seul le bon préréglage d'après les colonnes du fichier chargé.
   Dates : la valeur brute du classeur est lue (n° de série Excel → jj/mm/aaaa,
   SheetJS affichant sinon `m/d/yy`). Nombres : virgule, **au plus 2
   décimales**, sans espace ni unité ni « € » ; s'ils dépassent le `maxlength`
   du champ (6 pour conso/GES, 5 pour les estimations) → `VALEUR_TROP_LONGUE`
   (plus d'arrondi pour faire tenir). N° ADEME : trim + majuscules, format
   `^\d{4}[A-Z]\d{7}[A-Z]$`, sinon `ADEME_INVALIDE`. Une donnée obligatoire
   manquante, illisible ou invalide arrête la ligne **avant toute saisie**.

   **Déroulé (v2.16.0, en phases)**. Certains champs déclenchent un
   `__doPostBack` **complet** (hors de l'update panel `body_x_upcDiag`) : le
   formulaire est sérialisé au `submit()` puis la page entière est remplacée —
   tout ce qui est écrit ensuite dans l'ancien document est perdu. Donc :
   - **Détection des postbacks** : chaque document reçoit un `docId`
     (`crypto.randomUUID()`, repli `getRandomValues` sur http) posé sur
     `<html data-nh-doc>` par `content.js` au chargement (et renvoyé dans ses
     réponses) ou par la fonction injectée ; les événements `formdata`
     (synchrone au `submit()`), `beforeunload` et `pagehide` signalent une
     navigation en cours. **fireAndSettle** : `input` + `change` sur UN champ,
     puis interrogation de la frame jusqu'à un nouveau `docId` en
     `readyState complete` (rechargement) ou ~1,5 s de calme sans navigation
     en attente (pas de postback) ; 15 s max (réglable) → `POSTBACK_TIMEOUT`.
     Aucune référence d'élément n'est réutilisée après une attente : chaque
     action refait `getElementById`. Champs à postback détectés aussi
     automatiquement (`onchange` avec `__doPostBack`, `changeType`,
     `performanceEnergetique`, ou `input type=submit`).
   - **Phase A** — champs à postback, un par un : Type → Fournisseur →
     Établi le → Conso EP → GES. Pour chacun : relire le DOM, ne remplir que
     si la valeur normalisée diffère, fireAndSettle, vérifier qu'elle a tenu
     (3 essais). Sur la vraie page, un `change` synthétique sur « Établi le »
     n'atteint pas `performanceEnergetique` : après l'écriture de la date,
     `performanceEnergetique(el)` est appelé **directement dans le monde de
     la page** (`chrome.scripting.executeScript({ world: "MAIN" })`, même
     frame ; repli `__doPostBack(date)`), puis le rechargement est attendu et
     « Valable jusqu'au » comparé au calcul (journal `✓` / `✗`).
     « Valable jusqu'au » n'est **jamais** saisi : il doit valoir le calcul
     SIGEO (avant 2013 et dès le 01/07/2021 : +10 ans −1 jour ; 2013–2017 :
     31/12/2022 ; 01/2018–06/2021 : 31/12/2024), sinon « Établi le » est
     redéclenché.
   - **N° ADEME** (règle vérifiée le 23/11/2026) : en création, l'input
     `body_x_textIdDpe_x_textIdDpe` n'existe que si Type = DPE et « Établi
     le » ≥ 01/07/2021 (`ademeSaisissableAvantEnreg`). Dans ce cas il **doit**
     être là après le rechargement de la date, sinon `ADEME_CHAMP_ABSENT` et
     rien n'est enregistré. Pour un DPE plus ancien, il est saisi **après le
     1er enregistrement** : Enregistrer n°1 (id de la fiche lu dans l'URL ou
     `form.action` et journalisé avec le lien), saisie + relecture du N°,
     Enregistrer n°2. Une fois enregistré, le N° est un texte dans
     `span#body_x_labelIdDpe` (lu en priorité sur l'input). En simulation,
     le N° d'un DPE d'avant 07/2021 n'est pas tenté (indiqué au journal).
   - **Phase B** — ensuite seulement, les champs sans postback : N° ADEME
     (DPE ≥ 07/2021), date début validité, estimations,
     date de référence des prix, lettres forcées, cases, résultat,
     commentaire (valeur + `input`/`change`). **Lettres** : jamais comparées
     au PDF ; si la lettre du DPE diffère de celle calculée par SIGEO, elle
     est mise dans « Forcer la lettre » (option de l'étape, cochée par
     défaut) et c'est ce champ qui est vérifié.
   - **Phase C** — contrôle **avant** enregistrement : tout est relu et
     comparé (comparateur normalisé : `224` = `224,00`, `52,6` = `52,60`,
     écart < 0,005 ; dates jj/mm/aaaa ; ADEME en majuscules) ; les écarts
     sont ressaisis (2 essais) ; s'il en reste → `ECART_AVANT_ENREG` avec
     **tous** les écarts, rien n'est enregistré.
   - **Phase D** — **Enregistrer** (jamais « Enreg. et fermer » ni la touche
     Entrée) ; rechargement attendu (`docId`) et `Enregistrement réussi !`
     dans `#msgbody` (tout autre message arrête la ligne) ; contrôle
     **après** enregistrement avec le même comparateur (N° ADEME relu dans
     le champ ou dans `#body_x_labelIdDpe`) → `ECART_APRES_ENREG` avec tous
     les écarts, upload annulé.
   - **Phase E** — upload du PDF (postback attendu) puis vérification de la
     grille de `body_x_upcDiag` (plus de documents qu'avant l'upload **et**
     une ligne « DPE », relues dans le DOM). SIGEO n'affiche pas toujours le
     document tout de suite : si la vérification échoue, « Enregistrer » est
     recliqué sur la même fiche (attente active du rechargement, du bouton de
     nouveau cliquable et de la liste, 10 s max), puis la liste est revérifiée
     — **3 tentatives** au plus, **jamais de nouvel upload**. Document vu à
     une tentative : avertissement « document « DPE » apparu après le 2e
     nouvel enregistrement » et la ligne continue ; sinon « upload non vérifié
     après 3 tentatives d'enregistrement ». Une relance de cette ligne
     revérifie de même au lieu de ré-uploader (journal de session).
   - **Reprise** : tout rechargement inattendu fait repartir de la phase A
     d'après le DOM (aucun champ n'est sauté parce qu'il avait été « fait »).
     Un échec après l'Enregistrer n°1 est signalé « fiche DPE {id} créée mais
     incomplète (ADEME/PDF manquant…) : {lien} ». Une relance rouvre cette
     fiche en **modification** (journal de session, colonne résultat, ou URL
     qui contient déjà l'id) : jamais de second DPE.
   - **Journal** : pour chaque champ `attendu « … » · avant « … » · après « … »`
     et une ligne `↻ postback (champ)` à chaque rechargement.
   La popup reste ouverte. Champ ADEME absent alors qu'un N° est fourni →
   avertissement, la ligne continue.

   **PDF** : « PDF (ligne active) » (clic ou glisser-déposer) pour *Démarrer
   sur la ligne active* ; **Dossier des DPE** (mémorisé, sous-dossiers
   compris ; l'accès est redemandé au clic sur Démarrer) pour la boucle — le
   PDF « ligne active » n'est **jamais** utilisé en boucle. Correspondance
   stricte : nom exact (`fichier_dpe`, `.pdf` facultatif), sinon nom
   contenant le N° ADEME ; aucun fichier ou plusieurs → arrêt **avant toute
   saisie**. Contrôles : en-tête `%PDF` ; N° ADEME présent dans le nom, sinon
   dans le texte du PDF (pdf.js) ; un **autre** N° ADEME dans le nom ou le
   texte → arrêt ; aucun lisible (scan) → confirmation obligatoire (refus en
   boucle sans confirmation). Avant l'upload, le panneau affiche ligne, lot,
   fiche, N° ADEME (confirmé par SIGEO), fichier et contrôle, et demande
   confirmation selon « confirmer avant de joindre le PDF » (décoché : PDF
   joint automatiquement, ligne active comme boucle ; confirmation toujours
   demandée si le N° ADEME du PDF n'a pas pu être vérifié ; coché par
   défaut). N° ADEME « NC », « non communiqué », « non trouvé »,
   « [NON EMIS ADEME] », « 99999 »… : « NC » saisi dans le champ N° ADEME, avertissement
   (le PDF est alors retenu par son nom). Upload : Description « DPE » puis
   fichier posé dans le champ (`DataTransfer`, octets transmis en base64), le
   `change` déclenche l'envoi ; vérifié par le compteur « N Résultat(s) »
   (+1) et la ligne « DPE ».

   **Jamais de doublon** : une fiche déjà enregistrée (id dans l'URL ou
   N° ADEME affiché) n'est réenregistrée que si une valeur a été corrigée ;
   un document « DPE » déjà joint n'est pas renvoyé. Un journal
   (`chrome.storage.session`, écrit avant chaque postback) mémorise chaque DPE
   enregistré pendant la session du navigateur : relancé sur une page de
   création `/0`, l'étape rouvre la fiche existante (si l'URL de la fiche est
   renseignée) ou s'arrête ; un enregistrement lancé sans confirmation
   (rechargement en plein enregistrement) bloque la ligne jusqu'à
   vérification (« Vider l'historique DPE »). Les lignes dont la colonne
   résultat (`Saisie DPE` par défaut) commence par `OK` sont sautées.
   **Simulation** : tout est rempli et contrôlé, sans Enregistrer ni PDF.

   **Debug** : la case « debug » relève à chaque chargement tous les champs
   (id, libellé, visible / masqué, valeur, `onchange` / `onkeydown`…), les
   messages et la grille (résumé dans le journal, détail dans la console) ;
   « Lister les champs de la page » fait un relevé sans rien modifier,
   « Copier le relevé » copie le tout en JSON.

## 3. Onglet Extraction (ex-OSA)

**Type d'extraction** : *Standard* (ci-dessous), **Recherche IM SIGEO**,
**Recherche d'id interne de lot**, **DPE SIGEO**, **DPE extraction** ou
**Analyse DPE** (voir plus bas).

1. **Conditions** : ignorer certaines lignes (ex : colonne X = "NON").
2. **Champs de recherche** : sélecteur CSS (bouton 🎯) + colonne à taper.
3. **Validation** : touche Entrée ou clic sur un bouton, puis délai d'attente.
4. **Résultats à récupérer** : sélecteur CSS, ou « ligne de tableau » (compare
   une colonne du fichier à une cellule de chaque ligne du tableau de la page,
   puis extrait une autre cellule). Destination : colonne existante, lettre,
   ou **nouvelle colonne** créée à droite du tableau.
5. **Exécution** : plage de lignes, ▶ Démarrer / ■ Arrêter, journal et
   récapitulatif détaillé.

### Recherche IM SIGEO (adresse → N° E.I.)

Choisis **Recherche IM SIGEO** dans *Type d'extraction*, ouvre la recherche
d'immeubles SIGEO dans l'onglet cible (ou renseigne son URL), indique la
**colonne à chercher** (nom d'immeuble ou adresse, ex. « GRANDE ARMEE 62 »,
et éventuellement CP / ville), puis **Démarrer**.

Par défaut (*Nom : valeur complète, puis nom de voie*) : la valeur de la
cellule est tapée **telle quelle** dans le champ **Nom**, l'extension valide
(« Rechercher »), attend la fin du chargement et lit **toutes les pages** du
tableau. Une ligne correspond si son **Nom E.I.** est identique à la valeur,
ou si son **Adresse** correspond (voir ci-dessous). Si rien ne correspond,
elle relance la recherche avec le seul nom de voie (« ARMEE ») et refait la
comparaison sur les adresses.

Pourquoi ne pas se contenter du Nom : le champ **Nom** est un « contient »
sur un libellé libre, saisi de façon incohérente (`LAMARCK "112"`,
`14 LAMARCK`, `*98 LAMARCK***30/03/06****`) — le **Nom E.I.** ne sert donc
qu'en **égalité stricte** avec la valeur cherchée. La colonne **Adresse** est
normalisée (`<num>[ BIS|TER] <TYPE VOIE> <NOM VOIE> <CP> <VILLE>`), parfois
précédée du nom de l'immeuble (`grande Armee 62 62 AVENUE DE LA GRANDE ARMEE
75017 PARIS`) : ce préfixe est ignoré.

Pour chaque ligne :
1. **Parsing** de l'adresse : numéros (`158/160` → 158 et 160, `128 128BIS` →
   128 et 128BIS), suffixe BIS/TER, type de voie (RUE, BD, AV…), nom de
   voie, CP/ville (dans l'adresse ou dans des colonnes à part).
2. **Recherche** avec le seul **mot le plus distinctif** de la voie
   (`LAMARCK`, `SULPICE`, `CHARONNE`…) dans le champ Nom, **Site forcé à
   vide**. Variante : « Par adresse » (Adresse = ce mot, Code postal si
   connu), ou **les deux** — le journal indique alors combien de
   correspondances chaque mode apporte.
3. Attente de la fin du postback (partiel ou complet), puis **lecture de
   toutes les pages** (« Page suivante >> » jusqu'à désactivation) ; le total
   lu est comparé à « N Résultat(s) ».
4. **Filtrage** sur la colonne Adresse : numéro **et** suffixe exacts
   (128 ≠ 128 BIS), type de voie s'il est donné des deux côtés, voie,
   CP/ville si fournis. Plage `158/160` : l'adresse contenant les deux
   numéros est préférée.
5. **Résultat** : colonne *IM* (par défaut nouvelle colonne « N° E.I. »,
   remplie seulement si OK) et colonne *Statut* (« Résultat IM ») :
   `OK — IM0398636`, `Multiple — IM0038854, IM0069522` (doublons : **jamais**
   de choix arbitraire), `Introuvable — proches : …` (3 adresses les plus
   proches), `ERREUR — …`. Le récapitulatif compte les lignes « à vérifier ».
   Si la colonne IM est la colonne adresse elle-même, une adresse non
   résolue n'est jamais effacée.
6. **Cache** : une voie déjà cherchée n'est pas relancée ; le cache est vidé
   à chaque démarrage. L'estimation affiche le nombre de voies différentes.

Sur la vraie page, les blocs « Par désignation » (Nom, N° Siret), « Par
adresse » (Adresse, Code postal, Ville) et « Par mandat » (Site…) sont dans
**un seul formulaire** : tous les critères remplis se combinent. Chaque
recherche vide donc les critères de l'autre mode (v2.20.2 : Nom vidé pour
une recherche par adresse, Ville vidée dans les deux cas).
Les champs du formulaire (Nom — `#body_x_txtLabel_x_Label` —, Site,
Adresse, Code postal, Ville), le bouton « Rechercher », la bascule « Par
adresse » (s'il y en a une) et le tableau sont détectés automatiquement ; sinon, désigne-les avec 🎯
dans **Sélecteurs avancés**. **Tester sur la ligne de départ** lance la
recherche sans rien écrire et affiche ce qui a été trouvé sur la page.

### Recherche d'id interne de lot (IM + LT → id SIGEO)

Les fiches lot SIGEO ont l'URL `…/sig_imm/lot_manage/syn/{IM}/{ID_INTERNE}` :
l'id interne (ex. 28) n'a aucun rapport avec le numéro de lot (ex. 204). Ce
type d'extraction le retrouve pour chaque ligne.

1. Colle tes couples IM / LT depuis Excel dans l'onglet **Données** (IM avec
   ou sans « IM » et zéros devant : `IM0393528`, `0393528`, `393528` ; LT avec
   ou sans « LT » et zéros : `LT000204`, `204`).
2. Choisis **Recherche d'id interne de lot**, puis les colonnes **IM** et
   **LT** (devinées d'après les en-têtes : `IM`, `IM_format`, `lot`, `LT`…).
3. Sois connecté(e) à SIGEO dans l'onglet cible, puis **Tester sur la ligne
   de départ** et ▶ Démarrer.

**Méthodes** :
- *POST direct* (préférée) : depuis l'onglet SIGEO (même origine → cookies de
  session), un `GET` de `lot_browse/syn` récupère `__VIEWSTATE`,
  `__VIEWSTATEGENERATOR`, `__EVENTVALIDATION` et les autres champs, puis un
  `POST` renvoie le formulaire avec IM (7 chiffres), LT, Niv. vide, Statut et
  le bouton « Rechercher » du bloc *Par code* (`name=value` s'il s'agit d'un
  bouton submit, `__EVENTTARGET` s'il s'agit d'un lien `__doPostBack`). Le
  HTML reçu est lu avec `DOMParser`. L'onglet ne bouge pas.
- *Onglet* (repli) : la page `lot_browse` est ouverte, les champs remplis,
  « Rechercher » cliqué ; on attend la fin du postback (rechargement complet
  ou UpdatePanel) puis on lit les liens.
- *Auto* : POST direct ; si le serveur le refuse (HTTP 500, validation
  d'événement, champ LT non renvoyé…), bascule sur l'onglet pour la suite.

Dans les résultats, seuls les liens `a[href*="lot_manage"]` dont l'IM (dans
l'URL) et le numéro `LT000204` (dans le texte, espaces insécables
normalisés) correspondent numériquement sont retenus — la recherche SIGEO
étant un « contient », LT 204 renvoie aussi LT 2040. Écrit : **ID interne**,
**URL lot** (fiche) et **Résultat lot** : `OK — id 28`, `Introuvable`,
`Ambigu — id 40 (…), id 41 (…)` (jamais de choix arbitraire) ou
`ERREUR — …`. Si la page de connexion apparaît (session expirée),
l'extraction **s'arrête** et indique la ligne à partir de laquelle relancer.
Réglages : statut (`actif` par défaut, vide = valeur de la page), URL de
`lot_browse`, timeout, pause entre lignes ; sélecteurs avancés si les ids
SIGEO changent. Un même couple IM/LT n'est cherché qu'une fois par exécution.

### DPE SIGEO (vérifier + télécharger le DPE le plus récent par lot)

À lancer sur le fichier produit par la recherche d'id de lot
(`Immeuble | IM | lot | IM_format | ID interne | URL lot | Résultat lot`).

1. Connecte-toi à SIGEO dans l'onglet cible, charge le fichier (Données).
2. Choisis **DPE SIGEO**. **URL** : `{URL lot}` par défaut (la fiche lot ;
   le lien vers la page Documents `…/popup.aspx/fr/fil/ged_multi_file_manager/…`
   y est repéré et suivi). Si tu connais l'URL de la page Documents, mets-la
   directement avec des `{Colonne}`, ex. `…/lot_manage?…={ID interne}`.
3. **Tester sur la ligne de départ** : ouvre la page, liste les DPE trouvés et
   le nom de fichier prévu, sans rien télécharger ni écrire.
4. **Lancer** (permission « Téléchargements » déclarée dans le manifest).

Pour chaque ligne : attente de la page (« Chargement en cours... » disparu),
lecture des sections dont le **titre** contient `EPC` (« Diagnostics >> EPC »,
« Diagnostics Multys >> EPC » — repérées par leur texte, jamais par position
ni par id), filtre des fichiers par leur **nom** normalisé (sans accents,
« ? » d'encodage abîmé supprimés : `chaudie?re` → `chaudiere`) :
DPE si `dpe` / `diagnostic de performance énergétique` / `performance
énergétique`, sauf gaz, élec, chaudière, entretien, ENT, amiante, plomb,
termite, Carrez, ERP / ERNMT — sauf si « DPE » figure dans le nom. Les règles
(regex, une par ligne) sont modifiables dans « Règles de détection ».
**Le plus récent** : date de la colonne Date, puis année `20xx` du nom, puis
ordre du tableau.

Téléchargement **sans cliquer** sur le lien : le `href` `__doPostBack` du
lien n'est pas utilisé (il renvoie la page Documents). L'URL du fichier est
lue dans son `onclick="popup('/bare.aspx/fr/fil/ged_file_download/<ID>', …)"`
puis confiée au service worker, qui appelle `chrome.downloads.download`
(Chrome suit la redirection et envoie les cookies de session, ce qu'un
`fetch` depuis la page ne peut pas faire). Le service worker attend la fin :
interrompu → `KO téléchargement : téléchargement interrompu (…)` ; type
`text/html` → fichier supprimé et `KO téléchargement : SIGEO a renvoyé une
page HTML (session expirée ?)` ; lien sans `ged_file_download` →
`KO téléchargement : lien de téléchargement introuvable`. La boucle continue
dans tous les cas. Fichier enregistré dans `Téléchargements/DPE_SIGEO/`
(configurable) sous le nom `{IM}_{lot}_DPE_{date}` (`{nom}` = nom du fichier
dans SIGEO, sans extension) + extension du fichier SIGEO (`.pdf` par défaut),
date ISO (`IM0393664_10_DPE_2022-10-19.pdf`), caractères interdits Windows et
espaces retirés ; un doublon est renommé `(1)`.

Colonnes écrites (créées si absentes) : **DPE présent** (OUI / NON /
ERREUR), **Nb DPE** (plus d'un → ligne « à vérifier » dans le récap),
**DPE nom SIGEO**, **DPE date**, **DPE section**, **DPE source**
(gérance / syndic / gérance + syndic), **DPE fichier**, **DPE statut**.

**Repli syndic** (v2.17.3) : si la page gérance (`metier=ger`) ne contient
aucun DPE, ou est introuvable (URL vide, page absente), la même recherche est
faite dans la fiche lot syndic, par défaut
`http://sigeo.evoriel.net/page.aspx/fr/sig_imm/lot_manage/syn/{IM_format}/{lot}`
(ex. `…/syn/393151/90`) ; le lien vers la page Documents y est suivi même si
la fiche contient d'autres tableaux. Champ « Repli syndic » vide = pas de
repli. Pas de repli quand un DPE est trouvé en gérance (même si son
téléchargement échoue) ni quand la session est expirée. Le texte de DPE statut se règle dans « Texte écrit dans DPE
statut » : par défaut `OK & téléchargé`, `OK trouvé (non téléchargé)`
(vérification seule), `KO pas trouvé`, `KO téléchargement : {raison}`,
`Page introuvable`. Variables : `{nom}`, `{date}`, `{fichier}`, `{nb}`,
`{section}`, `{source}`, `{raison}` et toute `{Colonne}` de la ligne.
Session expirée → arrêt avec la ligne à partir de laquelle relancer.
Pause entre lots réglable (1 s par défaut) pour ne pas surcharger SIGEO.

### DPE extraction (immeuble + lot → fichier prêt pour la saisie SIGEO)

Option tout-en-un (ex-« DPE Intralia ») : **le fichier d'entrée n'a besoin
que de deux colonnes, le nom de l'immeuble et le n° de lot**. Pour chaque
ligne, dans l'ordre : Intralia (n° d'immeuble, page DPE du lot, DPE saisi,
DPE le plus récent de la GED) → FlowerDocs (téléchargement du PDF, lecture du
fournisseur / n° FU et du n° ADEME) → **SIGEO** (IM de l'immeuble retrouvé
par son nom sur le site SIGEO choisi, id interne du lot, URL de la fiche DPE) → une ligne du
fichier prêt pour la saisie. Logique dans `intralia-dpe.js`.

**Intralia n'est que lu** : uniquement des `GET` ; le formulaire du lot
(`form1`, qui enregistrerait le DPE) n'est jamais soumis, les cases `IDGED`
(rattachement de document) ne sont jamais cochées, « Passer la commande DPE »
n'est jamais cliqué.

1. Dans ce navigateur : connecte-toi à **Intralia**, à **FlowerDocs** et à
   **SIGEO**, et laisse ouverte dans un onglet SIGEO la **recherche
   d'immeubles** (page avec « Nom », « Site », « Rechercher ») — ou renseigne
   son URL dans *Recherche IM SIGEO*.
2. Charge le fichier (Données), choisis **DPE extraction** : les colonnes
   **immeuble** (`Immeuble`) et **lots** (le vrai n° de lot, jamais
   `IM_format`) sont devinées.
3. **Cabinet Intralia** : Sicov (19) par défaut ; la liste complète est lue
   dans le `<select name="ID_SOCIETE">` de `diag.php` (bouton ↻, ou
   automatiquement au premier chargement de l'index). *Tous les cabinets*
   (100) : ~90 Mo, ~2 min de lecture la première fois.
4. **Lancer**. **Tester sur la ligne de départ** fait toute la chaîne sur une
   ligne sans rien écrire ni télécharger.

Tout le reste est dans **Options** (déplié au besoin) : mode *télécharger*
(défaut) ou *repérage seul*, dossier et nom des PDF, analyse du PDF / OCR,
étape SIGEO (case « trouver la fiche DPE dans SIGEO »), colonne IM (défaut
*auto*) et colonne id interne, colonnes écrites (*saisie* par défaut ou
*toutes*), vérification du contenu, Toolbox, champs Intralia, réglages
avancés. Les réglages enregistrés avant la v2.20.0 sont remis sur le
tout-en-un (télécharger + analyse + SIGEO + colonnes de saisie).

**Étape 1 — index des immeubles** (une seule fois par traitement) :
`GET /gerance/diag.php?action=R&ID_SOCIETE={id}&ID_GEST=0&TYPEB=0&DPE_TOUS=checked&GES=0`
(`TYPEB=0` = tous les baux : garages, commerces… compris). La page est
rendue côté serveur (DataTables sans AJAX : tout est dans le HTML, la
« Recherche rapide » n'est pas utilisée), décodée en **windows-1252**
(`TextDecoder`, jamais `res.text()`), et les lignes de `#example3` sont lues
par **regex** (pas de `DOMParser` : 91 Mo pour tous les cabinets). Ids lus
dans le lien `diaglot.php?id_societe=…&id_immeuble=…&id_lot=…` de chaque
cellule. Index compact `NOM_NORMALISÉ|LOT` mis en cache dans
`chrome.storage.local` **par cabinet**, avec sa date ; relu au démarrage
s'il a plus de 7 jours (réglable, 0 = à chaque traitement) ou avec
**Rafraîchir l'index Intralia**.
Normalisation : nom sans accents, majuscules, tout non-alphanumérique →
espace (`Parc Fleuri - le Concerto 1` = `PARC FLEURI LE CONCERTO 1`) ; lot
sans zéros de tête (`0732` = `732`). Un même nom + lot trouvé plusieurs fois
(autre cabinet, ou deux immeubles homonymes) → `AMBIGU` (correspondances
listées, rien n'est téléchargé) ; nom trouvé sans le lot →
`LOT_INTROUVABLE` ; nom absent → `IMMEUBLE_INTROUVABLE` + 3 noms proches.

**Étape 2 — page du lot** :
`GET /gerance/diaglot.php?id_societe=…&id_immeuble=…&id_lot=…`. Le **DPE déjà
saisi** dans Intralia (formulaire `form1`) peut être recopié : cases à cocher
dans « DPE déjà saisi dans Intralia → colonnes » (lettre DPE, kWh/m², lettre
GES, kgCO2/m², type avant/après le 01/07/2021, `DATEDIAG`, fin de validité,
coûts min/max/estimé, année de référence, référence). Les **documents GED**
(`input.IDGED` → nom en gras, date « crée le AAAA-MM-JJ », UUID) contiennent
tous les rapports (chaudière, EDL, additifs, .msg…) : les DPE sont repérés
par leur **nom** (sans accents, insensible à la casse) avec la regex
modifiable `dpe|diag.*perf|perf.*energ|(^|[_\s.-])dp([_\s.-]|$)`
(`dpe.pdf`, `DUPONT_DPE.pdf`, `MARTIN_928_24.08.11_DP.pdf`…), dédoublonnés
(même nom + même date) ; **le plus récent par date « crée le »** est retenu
(à égalité, le premier de la page). Attention : « crée le » est la date de
dépôt dans la GED, pas celle du DPE — les deux sont exposées
(`DPE_Date_GED` et `DATEDIAG`). Aucun candidat → `AUCUN_DPE_GED` + nombre de
documents présents.
Option **vérifier le contenu** (désactivée par défaut) : les candidats sont
chargés du plus récent au plus ancien et le premier dont la page 1 contient
« performance énergétique » est retenu (texte du PDF via pdf.js, OCR local de
la Toolbox si la page est scannée).

**Étape 3 — téléchargement FlowerDocs** :
`GET https://www.oralia.flowerdocs.cloud/flower-docs-gui/download?documentId={UUID}&index=0`
(ce que fait le bouton « Télécharger » de FlowerDocs), avec la session du
navigateur — depuis le panneau, ou à défaut depuis un onglet FlowerDocs ouvert.
Le statut HTTP et les premiers octets (`%PDF`) sont **contrôlés avant tout
enregistrement** : page HTML / 401 / 403 → `FLOWERDOCS_NON_CONNECTE` et arrêt
de la boucle (une page de connexion n'est jamais enregistrée en .pdf).
Enregistrement par `chrome.downloads` (doublon renommé) dans
`Téléchargements/DPE/` sous le nom `{Immeuble}_lot{lots}_{DPE_Date_GED}.pdf`
(`{Immeuble}` et `{lots}` = les colonnes choisies, quel que soit leur nom
dans le fichier ; modifiable : `{Colonne}`, `{DPE_Date_GED}`, `{DPE_Fichier}`, `{DPE_UUID}`,
`{Intralia_ID_Immeuble}`, `{Intralia_Cabinet}`… ; caractères interdits par
Windows → `_` ; extension = celle du fichier reçu). Pour les étapes
`pdfcheck` / `pdfwrite` du scénario : cocher **ajouter les PDF téléchargés
à la Toolbox**, puis retrouver le PDF par `{DPE_Fichier_Local}`.

**Étape 4 — SIGEO : IM, id interne, URL de la fiche** (v2.20.0) :
- **IM** (v2.22.0) : colonne IM *auto* (défaut) → recherche d'immeubles
  SIGEO (`…/sig_imm/immeuble_browse/ger`) **par désignation + site** :
  - **Site SIGEO** (réglage de DPE extraction, enregistré avec les autres) :
    `391` (ORALIA SOGIMAT) par défaut ; la liste se remplit avec les sites lus
    sur la page. Code absent de la liste « Site » → arrêt avec « site X
    introuvable dans la liste ». Vide = tous les sites (déconseillé).
  - Avant chaque recherche : Adresse, Code postal, Ville, Siret, N° E.I.,
    codes copropriété / programme commercial vidés ; Filtre et Type sur
    « Tous » ; Nom = nom de l'immeuble du fichier ; Site = site choisi, avec
    son libellé écrit dans le champ caché `body_x_sites_x_hdSiteLabel`
    (`input` + `change` sur chaque champ modifié, sauf postback). Jamais
    « Réinit. ». Champs trouvés par leurs id exacts (`IM_GER_IDS` dans
    `sigeo-im.js`), sinon par leurs libellés.
  - Jamais d'`alert()` bloquante : pas de clic si Adresse est remplie sans
    CP ni ville, et toute alert/confirm pendant le clic est interceptée.
  - Le clic sur Rechercher recharge toute la page : c'est le panneau
    latéral qui pilote (il survit au rechargement), il attend la nouvelle
    grille `#body_x_grid_x` puis la lit (en-têtes, ou positions [1] N° E.I.,
    [2] Nom E.I., [3] Adresse ; id immeuble = fin du lien `a.edit_link`).
  - Choix (le champ Nom est un « contient ») : **exactement 1 Nom E.I.
    identique** (majuscules, sans accents, espaces réduits) → retenu ;
    sinon départage par le **CP puis la ville** de l'adresse Intralia
    (colonne Adresse) ; sinon rien → `À VÉRIFIER` avec le nombre de
    résultats et les candidats. Jamais la première ligne. Si le fichier ne
    donne rien, le nom Intralia est essayé (s'il diffère).
  - **Repli par adresse** (case, décochée par défaut) : l'ancienne recherche
    par l'adresse Intralia, toujours sur le même site, « par adresse »
    seulement avec le CP ; plusieurs IM → départagés par le lot.
  - Colonnes de contrôle : `ID immeuble SIGEO` (lien de la fiche immeuble)
    et `Méthode IM SIGEO` (ex. « désignation + site 391 — « LE BOURG » :
    3 résultats, 2 noms identiques, retenu par CP 69490 »).
  Une colonne IM peut aussi être donnée (`IM`, `IM SIGEO`, `IM_format`…,
  devinée) : remplie avec un IM → utilisée telle quelle ; vide ou sans IM
  (ex. un nom d'immeuble) → recherche auto (la recherche d'immeubles est
  alors nécessaire, comme sans colonne), et l'IM trouvé est écrit dans une
  case vide. Pour les immeubles introuvables, il suffit d'ajouter l'IM à la
  main dans cette colonne et de relancer.
- **Id interne du lot** : IM + n° de lot, avec les réglages et le cache de
  *Recherche d'id interne de lot* ; réutilisé s'il est déjà dans la colonne
  « id interne » (ex. `id_server`) et que c'est un nombre.
- (v2.21.1) Les listes « IM SIGEO déjà connu » et « id interne déjà connu »
  sont sur **auto** par défaut (en tête de liste). Les colonnes du nom
  d'immeuble et du n° de lot y sont refusées (remises sur auto, message),
  y compris dans un réglage ou un profil enregistré : avant, « IM SIGEO =
  immeuble » donnait « IM invalide « LE BOURG » » sans IM, id ni URL, et
  « id interne = lots » aurait pris le n° de lot pour l'id.
- **URL de la fiche DPE** :
  `http://sigeo.evoriel.net/popup.aspx/fr/sig_imm/diag_manage/{IM}/{id interne}/0`
  (modèle modifiable dans « Réglages avancés »).
- Même onglet SIGEO pour les deux recherches : celui où la recherche
  d'immeubles est ouverte (repéré par « Nom » + « Rechercher » + « Site » ou
  la bascule « Par nom / Par adresse »). Aucun onglet de ce type et pas d'URL
  réglée → arrêt avec un message clair. Session SIGEO expirée → arrêt avec la
  ligne de reprise. Échec pour une ligne → « URL fiche DPE » vide, `DPE
  analyse` = À VÉRIFIER et la raison dans `Remarques DPE`.
Le fichier final (mode par défaut « fichier prêt pour la saisie SIGEO ») ne
contient que : `DPE analyse`, `URL fiche DPE`, `Type de local` (type du lot
Intralia : appartement / maison… → habitation, commerce / bureau / local… →
hors habitation, autre → à vérifier), `PDF analysé`, `N° FU`,
`Fournisseur DPE`, `Établi le`, `Valable jusqu'au`, `N° ADEME`, énergie
primaire, classe énergie, GES, classe GES, coûts min/max, `Date d'indexation
des prix` (DPE ≥ 01/07/2021 : 01/01/année de référence Intralia, sinon le
PDF), `Remarques DPE`, précédées de `IM SIGEO` (IM + 7 chiffres, si IM
auto), `ID immeuble SIGEO`, `ID interne SIGEO` (du lot) et `Méthode IM SIGEO`. Ensuite : onglet Saisie →
**Modèle DPE** reconnaît ce fichier et règle l'étape « SIGEO : saisir un DPE »
toute seule (préréglage « Fichier DPE extraction » : colonnes + URL
`{URL fiche DPE}` reprise telle quelle) ; il reste à choisir le dossier
`Téléchargements\DPE` et à lancer le scénario (en simulation d'abord).

**Colonnes pour la saisie SIGEO** (v2.18.2) : mêmes noms que l'Analyse DPE,
lus par l'étape « SIGEO : saisir un DPE » (préréglage « Colonnes de l'Analyse
DPE ») — `DPE analyse` (OK / À VÉRIFIER / Pas de PDF / ÉCHEC), `PDF analysé`,
`N° FU`, `Fournisseur DPE`, `Établi le`, `Valable jusqu'au`, `N° ADEME`,
`Énergie primaire (kWh/m²/an)`, `Classe énergie`, `GES (kg CO2/m²/an)`,
`Classe GES`, `Coût min/max (€/an)`, `Remarques DPE`.
- **Données du DPE = celles saisies dans Intralia** (`DATEDIAG`, kWh/m²,
  lettres, kgCO2/m², coûts min/max pour un DPE ≥ 01/07/2021). « Valable
  jusqu'au » est calculé comme SIGEO (la `datefinval` d'Intralia n'est pas
  fiable). Coûts : les cases « Montant théorique des dépenses énergétiques :
  Entre … et … » de la page du lot. Champ vide ou à 0 dans Intralia → valeur
  lue dans le PDF, signalée (v2.18.3).
- **Fournisseur (→ N° FU) et N° ADEME = lus dans le PDF** (case « analyser le
  PDF », cochée par défaut ; moteur de l'Analyse DPE, texte + OCR local
  réglable). Le n° FU vient de la liste « Fournisseurs → n° FU » de l'Analyse
  DPE ; fournisseurs inconnus signalés en fin de traitement. Pour un DPE
  d'avant le 01/07/2021, pas de remarque sur le n° ADEME (pas de champ en
  création dans SIGEO). Date du PDF ≠ Intralia → à vérifier.
- `Remarques DPE` ne garde que ce qui demande une action ; sans DPE retenu,
  elle reprend le statut (`AUCUN_DPE_GED — …`).
- **Colonnes écrites** : *utiles à la saisie SIGEO seulement* (par défaut)
  ou *toutes* (détail Intralia, liens, UUID… ci-dessous).
En *repérage seul*, le PDF est lu en mémoire sans être enregistré.

En mode « toutes », colonnes écrites en plus : `Intralia_Cabinet`,
`Intralia_ID_Immeuble`, `Intralia_URL_Lot`, `DPE_Statut` (`OK` |
`AUCUN_DPE_GED` | `LOT_INTROUVABLE` | `IMMEUBLE_INTROUVABLE` | `AMBIGU` |
`INTRALIA_NON_CONNECTE` | `FLOWERDOCS_NON_CONNECTE` | `ERREUR`),
`DPE_Detail` (correspondances, noms proches, nombre de documents, autres
DPE…), `DPE_Fichier` (nom dans la GED), `DPE_Date_GED`, `DPE_UUID`,
`DPE_Lien_FlowerDocs`, `DPE_Fichier_Local` + les champs cochés
(`Intralia_DPE_Classe`, `Intralia_DPE_kWh_m2`, `Intralia_GES_Classe`,
`Intralia_GES_kgCO2_m2`, `Intralia_Date_Diag`, …) + les colonnes de l'analyse
du PDF (ci-dessus).
Traitement séquentiel, pause de 300 à 500 ms entre deux lots (réglable),
bouton ■ Arrêter, journal par ligne ; un même nom + lot n'est traité qu'une
fois. Page de connexion Intralia → `INTRALIA_NON_CONNECTE`, arrêt avec la
ligne à partir de laquelle relancer. URL Intralia / FlowerDocs modifiables
dans « Réglages avancés ».

### Analyse DPE (dossier de PDF → Excel)

Lit les PDF de DPE téléchargés (par ex. `Téléchargements\DPE_SIGEO`, rempli
par *DPE SIGEO*) et écrit, pour chaque lot, les informations de la saisie
SIGEO + le **n° fournisseur FU**. Tout est lu sur le PC : pdf.js pour la
couche texte, Tesseract.js (OCR local) pour les PDF scannés, les images et
l'étiquette énergie. Logique dans `dpe-scan.js` (port de l'outil Rust
*DPE Scan*).

1. Charge le fichier produit par *DPE SIGEO* (Données) — ou rien : une ligne
   est alors créée par PDF (`DPE fichier | IM | lot`, IM et lot lus dans le
   nom `IM0393664_10_DPE_2022-10-19.pdf`).
2. Choisis **Analyse DPE**, puis **Choisir le dossier** (sous-dossiers
   compris ; PDF, JPG, PNG). Les fichiers ne sont pas conservés à la
   fermeture du panneau.
3. **Rattacher par** : colonne `DPE fichier` (nom exact), sinon **IM + lot**
   lus dans le nom du PDF (`IM0393664`, `0393664`, `393664` ; `LT000010`,
   `10`). Deux PDF pour le même lot : le plus récent (date du nom) est
   analysé, l'autre signalé. PDF sans ligne : ajoutés en bas du tableau
   (option, seulement si la plage va jusqu'à la fin). Les colonnes sont
   devinées d'après les en-têtes.
4. **OCR** : *auto* (pages scannées, et page DPE si une valeur manque),
   *toujours la page DPE* (recoupe l'étiquette, plus lent) ou *jamais*.
5. **Tester sur la ligne de départ** (rien n'est écrit), puis ▶ Démarrer, puis
   **Télécharger l'Excel** (même export que l'onglet Export).

Pour chaque PDF : repérage de la page DPE (score : « établi le »,
« valable jusqu'au », n° ADEME, « estimation des coûts annuels »… — même au
milieu d'un dossier de diagnostics) et de la page des consommations (p.3) ;
dates (cohérence établi + 10 ans) ; n° ADEME (13 caractères, réparé si lu par
OCR) ; surface ; **énergie primaire et GES par m² = plancher(total ÷ surface)**
(page 3 et « ce logement émet … kg de CO₂ »), recoupés avec l'étiquette lue
par OCR quand elle l'est ; classes calculées (double seuil 2021, barème des
petites surfaces < 40 m² pour les DPE établis depuis le 01/07/2024, ancien
barème avant juillet 2021) ; fourchette de coûts ; société, diagnostiqueur,
n° de certification, adresse du bien. Anciens DPE (2006 → juin 2021) :
« consommation conventionnelle », « estimation des émissions », validité
légale rappelée (31/12/2022 ou 31/12/2024).

**Fournisseurs → n° FU** (bloc repliable) : une ligne `NOM : numéro`
(aussi `->`, `;`, tabulation), enregistrée dans le navigateur et commune à
tous les profils ; import / export `.txt` (UTF-8 ou Windows-1252). Le nom est
cherché dans la société lue, puis le bloc diagnostiqueur (e-mail compris,
`contact@acade.fr` → ACADE), puis le texte du DPE ; casse, accents et
ponctuation ignorés (`ACT'IM`, `Arthemis Côte d'Or`), noms de personne dans
le désordre acceptés (`Frédérique ROUSSEL`). Faute d'une lettre sur un nom
de 6 caractères ou plus (sans chiffre) : rapproché mais « À VÉRIFIER »
(`SOCOTEC` ≈ `SOCOTECH`). Les sociétés absentes de la liste sont listées à la
fin ; **Ajouter les fournisseurs inconnus à la liste** les met en bas de la
liste, numéro à compléter.

Colonnes écrites (créées si absentes) — **uniquement ce qui sert à la saisie
du DPE dans SIGEO** (v2.17.2) : `DPE analyse` (OK / À VÉRIFIER / ÉCHEC / Pas
de PDF / PDF introuvable), `PDF analysé` (fichier à joindre), `N° FU`,
`Établi le`, `Valable jusqu'au`, `N° ADEME`, `Énergie primaire (kWh/m²/an)`,
`Classe énergie`, `GES (kg CO2/m²/an)`, `Classe GES`, `Coût min (€/an)` et
`Coût max (€/an)` (ces deux-là seulement pour un DPE établi depuis le
01/07/2021, SIGEO n'ayant pas d'estimation avant), et `Remarques DPE` (ce qui a fait passer la ligne
« À VÉRIFIER » : valeur manquante, lue par OCR seule, écart étiquette /
calcul, fournisseur rapproché ou absent de la liste — avec le nom lu…).
Les totaux annuels, la surface, l'adresse, le diagnostiqueur et la
certification restent lus pour les recoupements mais ne sont plus écrits.
Le bouton « Colonnes de l'Analyse DPE » de l'étape de saisie utilise
`{PDF analysé}` pour le fichier ; la date d'indexation des prix est
facultative (colonne `date_prix` vide par défaut, sans avertissement).

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

Ouvre `test.html` dans Chrome (fichier local, non publié sur GitHub) : la page
contient un formulaire complet pour tester la Saisie (avec IBAN découpé) et
une recherche simulée avec tableau de résultats pour tester l'Extraction.
Pour la recherche IM, ouvre `tests/fixtures/sigeo-im-mock.html`.
Récap/reprise de la boucle : `NODE_PATH=$(npm root -g) node tests/scenario-recap.panel.e2e.js`.

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
├── sigeo-im.js        Extraction « Recherche IM SIGEO » (parsing, recherche, pagination, filtrage)
├── sigeo-lot.js       Extraction « Recherche d'id interne de lot » (POST direct, repli onglet)
├── sigeo-dpe.js       Extraction « DPE SIGEO » (page Documents, filtre, plus récent, téléchargement)
├── intralia-dpe.js    Extraction « DPE extraction » (index diag.php, page du lot, DPE GED le plus récent, FlowerDocs, fichier prêt)
├── dpe-scan.js        Extraction « Analyse DPE » (lecture des DPE, recoupements, classes, fournisseurs FU)
├── sigeo-diag.js      Étape de scénario « SIGEO : saisir un DPE » (formulaire diag_manage + PDF)
├── update-check.js    Vérification de mise à jour via GitHub
├── lib/xlsx.full.min.js  SheetJS (lecture/écriture Excel)
├── lib/pdf.min.js     pdf.js (Mozilla) — analyse PDF de la Toolbox
├── lib/pdf.worker.min.js  Worker pdf.js
├── lib/ocr/           Tesseract.js (OCR local) : lib, worker, cœur wasm, langue fra
├── icons/             Icônes de l'extension
├── test.html          Page de test locale des deux modes (non publiée)
└── tests/             Tests des recherches IM, lot et DPE (SIGEO, Intralia), analyse DPE, saisie DPE (sites simulés, DPE fictifs, jsdom / Playwright)
```

Tests de la recherche IM (Node + Playwright + Chrome) :
`node tests/sigeo-im.e2e.js` (logique + page simulée en postback partiel
et complet) et `node tests/sigeo-im.panel.e2e.js` (onglet Extraction du
vrai panneau). Ouvrir `tests/fixtures/sigeo-im-mock.html` dans Chrome
permet aussi d'essayer la recherche à la main.
Recherche de lot : `node tests/sigeo-lot.e2e.js` (POST direct, repli onglet,
ambigu / introuvable, session expirée, contre un faux serveur ASP.NET :
`tests/fixtures/sigeo-lot-server.js`) et `node tests/sigeo-lot.panel.e2e.js`.
DPE : `node tests/sigeo-dpe.jsdom.js` (jsdom : filtre, plus récent, nom de
fichier, attente du chargement, fiche → page Documents, POST → PDF ou HTML,
contre `tests/fixtures/sigeo-dpe-mock.js`), `node tests/sigeo-dpe.background.js`
(service worker : PDF, page HTML supprimée, interruption) et
`node tests/sigeo-dpe.panel.e2e.js`.
DPE extraction : `node tests/intralia-dpe.jsdom.js` (jsdom : décodage
windows-1252, index par regex — 36 000 lignes —, normalisation, collisions,
filtre DPE, tri « crée le », pages de connexion, contrôle du fichier
FlowerDocs, nom de fichier, 6 cas réels anonymisés, analyse des PDF par le
moteur Analyse DPE — n° FU, n° ADEME —, contre `tests/fixtures/intralia-mock.js`)
et `node tests/intralia-dpe.panel.e2e.js` (vrai panneau : repérage,
téléchargement, analyse, vérification du contenu par pdf.js, Toolbox,
sessions expirées, cache de l'index, profils, fichier d'entrée « immeuble +
lots » seulement avec IM retrouvé par l'adresse dans un faux SIGEO —
recherche d'immeubles + lot_browse sur la même origine, deux IM à la même
adresse départagés par le lot —, et depuis 2.22.0 la page
`immeuble_browse/ger` fidèle (`tests/fixtures/sigeo-im-ger.js` : postback
complet, id exacts, Site + libellé caché, alert si Adresse sans CP ni ville,
grille `#body_x_grid_x`) : LE BOURG → IM0399982 / fiche 399982 sur le site
391 malgré un homonyme du site 390, homonymes à départager, formulaire
laissé rempli, site absent, aucun POST vers Intralia).
Analyse DPE : `node tests/dpe-scan.node.js` (moteur : classes, n° ADEME,
fournisseurs, rattachement, puis 5 DPE fictifs — texte, scanné, ancien,
fournisseur inconnu, PDF sans ligne — générés par
`tests/fixtures/make_dpe.py` ; OCR par `tesseract` en ligne de commande) et
`node tests/dpe-scan.panel.e2e.js` (vrai panneau servi en HTTP local, choix du
dossier, OCR Tesseract.js, liste fournisseurs, export Excel). Un dossier de
vrais DPE peut être passé en argument : `node tests/dpe-scan.node.js <dossier>`.
Saisie DPE : `node tests/sigeo-diag.node.js` (logique : dates, nombres,
comparateur normalisé, calcul de « Valable jusqu'au », plan des phases,
lettres forcées, fireAndSettle à horloge virtuelle, repli du handler,
correspondance des fichiers, contrôle ADEME),
`node tests/sigeo-diag.panel.e2e.js` (vrai panneau contre un faux SIGEO
`diag_manage` — `tests/fixtures/sigeo-diag-server.js` : postbacks complets
sérialisés au `submit()`, fournisseur à postback, change synthétique sans
effet sur « Établi le », ADEME `body_x_textIdDpe_x_textIdDpe`, coûts après
la conso, refus serveur, upload multipart — : saisie de bout en bout, ancien
DPE, fichiers absent / ambigu / mauvais ADEME / scan / non PDF,
rechargements pendant la saisie et la confirmation, relances sans doublon)
et `node tests/sigeo-diag.realcase.e2e.js` (cas réel 448957 / 08/10/2020 /
2021L2000833E / 224 / 52,6 en simulation puis enregistré, serveur lent,
`ADEME_INVALIDE`, `VALEUR_TROP_LONGUE`, `ECART_AVANT_ENREG`,
`ECART_APRES_ENREG`, `POSTBACK_TIMEOUT`, lettre forcée ; depuis 2.17.0 :
ancien DPE 24/11/2020 + ADEME 2021N1001671X → 2 enregistrements, nouveau DPE
→ 1 seul, `ADEME_CHAMP_ABSENT`, fiche incomplète puis reprise sans doublon
par le journal et par la colonne résultat ; depuis 2.17.1 : PDF visible
seulement après un nouvel « Enregistrer », PDF jamais visible puis relances
sans nouvel upload).

**Version** : 2.22.0 — **DPE extraction** : IM SIGEO cherché **par désignation + site** (Site SIGEO réglable, 391 par défaut) au lieu de l'adresse sans site : tous les autres critères vidés, libellé caché du site écrit, aucune alert() bloquante ; 1 Nom E.I. identique → retenu, sinon CP / ville Intralia, sinon À VÉRIFIER ; repli par adresse facultatif (même site + CP) ; colonnes `ID immeuble SIGEO` et `Méthode IM SIGEO`.
2.21.1 — **DPE extraction** : IM et id interne toujours cherchés (Recherche IM SIGEO puis Recherche d'id interne de lot) quand « IM SIGEO » / « id interne » pointent sur la colonne immeuble / lots ou sur une case sans IM ; listes renommées « … déjà connu », « auto » en tête.
2.17.7 — **Saisir un DPE** : N° ADEME « NC » (ou non communiqué / non trouvé / 99999…) → « NC » est saisi dans le champ N° ADEME (après le 1er enregistrement pour un DPE d'avant 07/2021) ; le PDF est retrouvé par son nom ; une création que SIGEO n'a pas faite (page revenue en création) ne bloque plus la relance par un faux « enregistrement lancé sans confirmation ».
2.17.6 — **Saisir un DPE** : fournisseur — préfixe « FU » et espaces retirés avant la saisie, validation par perte du focus (jamais Entrée) ; contrôle sans zéros de tête (SIGEO stocke 0448957 comme 448957 : info au journal, pas d'écart) ET libellé `#body:x:proxyFou:x:lbl:x:lbl` non vide, sinon écart « Fournisseur non reconnu par SIGEO » ; reprise d'une fiche existante aussi par un lien `…/diag_manage/{imm}/{n}/{id}` ou « diagnostic {id} » dans la colonne résultat (champs vérifiés, puis upload direct).
2.17.5 — **Analyse DPE** : colonne `Prix indexés au` retirée (inutile à la saisie) ; dans l'étape de saisie, la date d'indexation des prix devient facultative.
2.17.4 — **Saisir un DPE** : N° ADEME « NC » / non communiqué / non trouvé / 99999 accepté (rien de saisi, avertissement) ; option « confirmer avant de joindre le PDF » valable aussi sur la ligne active (décochée → PDF joint automatiquement). **Analyse DPE** : N° ADEME absent → « NC » (évaluation thermique / non émis ADEME) ou « non trouvé ».
2.17.3 — **DPE SIGEO** : si aucun DPE n'est trouvé en gérance (ou si la page est introuvable), la fiche lot **syndic** est vérifiée (`…/page.aspx/fr/sig_imm/lot_manage/syn/{IM_format}/{lot}`, champ « Repli syndic », vide = désactivé) ; nouvelle colonne `DPE source` (gérance / syndic / gérance + syndic), variable `{source}` dans les textes de statut ; la page Documents est suivie depuis une fiche lot qui contient d'autres tableaux.
2.17.2 — **Analyse DPE** : n'écrit plus que les colonnes utiles à la saisie SIGEO (plus de totaux annuels, surface, adresse, diagnostiqueur, certification, lecture, fournisseur lu / de la liste, type) ; nouvelle colonne `Prix indexés au` ; coûts et date des prix seulement pour un DPE ≥ 01/07/2021 ; préréglage de l'étape de saisie : `{PDF analysé}`, `{Prix indexés au}`.
2.17.1 — **SIGEO : saisir un DPE** : faux négatif de la
vérification d'upload corrigé — PDF absent de la liste juste après l'upload
→ « Enregistrer » recliqué et liste revérifiée (3 tentatives, 10 s d'attente
active chacune), jamais de nouvel upload ; avertissement si le document
n'apparaît qu'après un nouvel enregistrement ; une relance revérifie au lieu
de ré-uploader.
2.17.0 — **SIGEO : saisir un DPE** : « Établi le » déclenché
par `performanceEnergetique(el)` dans le monde de la page (repli
`__doPostBack`) ; N° ADEME saisissable avant enregistrement seulement si le
DPE est établi depuis le 01/07/2021, sinon **enregistrement en deux temps**
(Enregistrer n°1, N° ADEME, Enregistrer n°2) ; `ADEME_CHAMP_ABSENT` au lieu
d'un avertissement ; N° relu dans `span#body_x_labelIdDpe` en priorité ;
« fiche DPE {id} créée mais incomplète » + reprise en modification (journal,
colonne résultat, URL avec id), jamais de second DPE.
2.16.0 — **SIGEO : saisir un DPE** fiabilisée face aux
postbacks complets : `docId` par document (content.js + fonction injectée),
fireAndSettle (rechargement ou 1,5 s de calme, 15 s → `POSTBACK_TIMEOUT`),
repli par appel direct du handler sur « Établi le », remplissage en phases
A (postback) → B (sans postback) → C (contrôle avant, 2 essais) →
D (Enregistrer + contrôle après) → E (PDF), N° ADEME par son id réel,
« Valable jusqu'au » vérifié mais jamais saisi, lettre du DPE forcée
seulement si SIGEO en calcule une autre, codes `ECART_AVANT_ENREG`,
`ECART_APRES_ENREG`, `VALEUR_TROP_LONGUE`, `ADEME_INVALIDE`, tous les écarts
listés.
2.15.0 — nouvelle étape de scénario **SIGEO : saisir un DPE**
(formulaire diagnostic + PDF joint, reprise après chaque postback, journal
anti-doublon, correspondance stricte du fichier, mode debug).
2.14.0 — nouveau type d'extraction **Analyse DPE** (dossier de
PDF → colonnes DPE + n° fournisseur FU d'après une liste modifiable).
2.13.3 — DPE SIGEO : textes de la colonne DPE statut
personnalisables (défaut « OK & téléchargé » / « KO pas trouvé »).
2.13.2 — DPE SIGEO : téléchargement par l'URL
`ged_file_download` du `onclick` popup(…) via `chrome.downloads` dans le
service worker (le POST `__doPostBack` renvoyait la page Documents) ;
permission `downloads` dans le manifest ; variable `{nom}`.
2.13.1 — DPE SIGEO : titres de section lus dans les
`<fieldset><legend>` de la page Documents (le formulaire SIGEO porte un
`onclick` qui masquait tous les titres en 2.13.0).
2.13.0 — nouveau type d'extraction **DPE SIGEO** (DPE le plus
récent de chaque lot : vérification, téléchargement renommé, colonnes DPE).
2.12.0 — nouveau type d'extraction **Recherche d'id interne de
lot** (IM + LT → id interne SIGEO + URL de la fiche, par POST direct avec
repli sur l'onglet).
2.11.2 — Recherche IM : la valeur complète est d'abord tapée
dans Nom (correspondance sur le Nom E.I. exact ou l'adresse), puis le nom de
voie si rien n'est trouvé ; adresses SIGEO préfixées du nom reconnues.
2.11.1 — la **Recherche IM SIGEO** passe dans l'onglet
**Extraction** (type d'extraction dédié) au lieu du scénario de Saisie.
2.11.0 — recherche IM SIGEO (adresse → N° E.I. : recherche par nom de voie,
toutes les pages, correspondance sur la colonne Adresse, statuts OK /
Multiple / Introuvable, cache par voie).
2.8.0 — scénario : nouvelle étape **Saisir un champ** (un seul
champ name/id/classe, valeur fixe ou `{Colonne}`, option « puis appuyer sur
Entrée ») ; l'étape **Cliquer sur un élément** peut **appuyer sur Entrée** au
lieu de cliquer (valide un formulaire comme au clavier) ; chaque étape affiche
une **icône colorée** selon son type (horloge pour Attendre, etc.).
2.5.0 — étape de scénario **SIGEO : saisir une adresse**
(formulaire address_manage : résolution code commune via autocomplétion ou
sélecteur, n° de voie par libellé, dryRun, rapport par ligne).
2.2.0 — Toolbox : **OCR local** (Tesseract) pour PDF scannés et
images (étiquette énergie DPE : conso kWh/m²/an, émissions kgCO₂/m²/an).
2.1.0 — onglet **Toolbox (bêta)** : analyse de PDF (n° ADEME, SIRET,
noms, dates…), vérification/report des données dans le scénario de saisie.
