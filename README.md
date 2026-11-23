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

## 3. Onglet Extraction (ex-OSA)

**Type d'extraction** : *Standard* (ci-dessous), **Recherche IM SIGEO**,
**Recherche d'id interne de lot**, **DPE SIGEO** ou **Analyse DPE** (voir
plus bas).

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

Les champs du formulaire (Nom — `#body_x_txtLabel_x_Label` —, Site,
Adresse, Code postal), le bouton « Rechercher », la bascule « Par adresse »
et le tableau sont détectés automatiquement ; sinon, désigne-les avec 🎯
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
**DPE nom SIGEO**, **DPE date**, **DPE section**, **DPE fichier**,
**DPE statut**. Le texte de DPE statut se règle dans « Texte écrit dans DPE
statut » : par défaut `OK & téléchargé`, `OK trouvé (non téléchargé)`
(vérification seule), `KO pas trouvé`, `KO téléchargement : {raison}`,
`Page introuvable`. Variables : `{nom}`, `{date}`, `{fichier}`, `{nb}`,
`{section}`, `{raison}` et toute `{Colonne}` de la ligne.
Session expirée → arrêt avec la ligne à partir de laquelle relancer.
Pause entre lots réglable (1 s par défaut) pour ne pas surcharger SIGEO.

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

Colonnes écrites (créées si absentes) : `DPE analyse` (OK / À VÉRIFIER /
ÉCHEC / Pas de PDF / PDF introuvable), `PDF analysé`, `Fournisseur DPE`,
`Fournisseur (liste)`, `N° FU`, `Type diagnostic` (Performance
Energetique), `Établi le`, `Valable jusqu'au`, `Énergie primaire
(kWh/m²/an)`, `Énergie primaire totale (kWh/an)`, `Classe énergie`, `GES (kg
CO2/m²/an)`, `GES total (kg CO2/an)`, `Classe GES`, `Coût min (€/an)`, `Coût
max (€/an)`, `N° ADEME`, `Surface (m²)`, `Adresse DPE`, `Diagnostiqueur`, `N°
certification`, `Lecture` (texte / OCR, page) et `Remarques DPE` (ce qui a
fait passer la ligne « À VÉRIFIER » : valeur manquante, lue par OCR seule,
écart étiquette / calcul, fournisseur rapproché ou absent…).

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
├── dpe-scan.js        Extraction « Analyse DPE » (lecture des DPE, recoupements, classes, fournisseurs FU)
├── update-check.js    Vérification de mise à jour via GitHub
├── lib/xlsx.full.min.js  SheetJS (lecture/écriture Excel)
├── lib/pdf.min.js     pdf.js (Mozilla) — analyse PDF de la Toolbox
├── lib/pdf.worker.min.js  Worker pdf.js
├── lib/ocr/           Tesseract.js (OCR local) : lib, worker, cœur wasm, langue fra
├── icons/             Icônes de l'extension
├── test.html          Page de test locale des deux modes (non publiée)
└── tests/             Tests des recherches IM, lot et DPE, analyse DPE (SIGEO simulé, DPE fictifs, jsdom / Playwright)
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
Analyse DPE : `node tests/dpe-scan.node.js` (moteur : classes, n° ADEME,
fournisseurs, rattachement, puis 5 DPE fictifs — texte, scanné, ancien,
fournisseur inconnu, PDF sans ligne — générés par
`tests/fixtures/make_dpe.py` ; OCR par `tesseract` en ligne de commande) et
`node tests/dpe-scan.panel.e2e.js` (vrai panneau servi en HTTP local, choix du
dossier, OCR Tesseract.js, liste fournisseurs, export Excel). Un dossier de
vrais DPE peut être passé en argument : `node tests/dpe-scan.node.js <dossier>`.

**Version** : 2.14.0 — nouveau type d'extraction **Analyse DPE** (dossier de
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
