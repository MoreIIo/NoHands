# Guide de démarrage rapide - NoHands

## Installation dans Chrome (2 minutes)

### 1. Ouvrir Chrome Extensions
1. Ouvrez Google Chrome
2. Tapez dans la barre d'adresse : `chrome://extensions/`
3. Appuyez sur Entrée

### 2. Activer le mode développeur
- Activez le toggle **"Mode développeur"** en haut à droite de la page

### 3. Charger l'extension
1. Cliquez sur **"Charger l'extension non empaquetée"**
2. Sélectionnez le dossier : `C:\Users\Lulu\Desktop\NoHands`
3. Cliquez sur **"Sélectionner un dossier"**

### 4. Épingler l'extension (optionnel)
1. Cliquez sur l'icône d'extension (puzzle) dans la barre d'outils Chrome
2. Trouvez "NoHands - Excel Data Formatter"
3. Cliquez sur l'icône d'épingle pour la garder visible

## Test rapide

### Données de test
Copiez cette ligne (avec les tabulations) :

```
12345	M.	DUPONT	Jean	12 Rue de la Paix	Appt 3	75001	Paris	0123456789	0123456790	0612345678	jean.dupont@email.com	FR7612345678901234567890123	Mensuel	Trimestriel	Virement	5%	O	2%	3%	O	Caution	2024-01-15	MARTIN	Sophie
```

### Utilisation
1. **Cliquez sur l'icône NoHands** (NH sur fond violet) dans la barre d'outils
2. **Collez les données** dans la zone de texte (Ctrl+V)
3. **Cliquez sur COPIER** (ou Ctrl+Enter)
4. **Vérifiez le message** "✓ Données copiées avec succès !"
5. **Collez quelque part** (Ctrl+V dans un éditeur de texte) pour voir le résultat

### Résultat attendu
```
N° PROP (TW): 12345
CIVILITE PROP: M.
NOM PROP: DUPONT
PRENOM PROP: Jean
ADRESSE LIGNE 1 PROP: 12 Rue de la Paix
ADRESSE LIGNE 2 PROP: Appt 3
CP PROP: 75001
VILLE PROP: Paris
TELEPHONE DOMICILE PROP: 0123456789
...
```

## Utilisation avec vos vraies données Excel

1. Ouvrez votre fichier Excel
2. Sélectionnez **toute la ligne** avec les 25 colonnes (de N° PROP à PRENOM GESTIONNAIRE)
3. Copiez (Ctrl+C)
4. Ouvrez l'extension NoHands
5. Collez dans la zone de texte
6. Cliquez sur COPIER
7. Collez dans votre formulaire web

## Dépannage rapide

### L'extension ne s'affiche pas
- Vérifiez qu'elle est activée dans `chrome://extensions/`
- Rechargez l'extension (icône de rafraîchissement)

### Erreur "Format invalide"
- Vérifiez que vous avez **exactement 25 colonnes**
- Les colonnes doivent être séparées par des **tabulations** (depuis Excel)
- Ne copiez **qu'une seule ligne** à la fois

### Le bouton reste grisé
- La zone de texte doit contenir du texte
- Vérifiez que vous avez bien collé les données

## Documentation complète
Voir [README.md](README.md) pour plus de détails sur :
- Les 25 colonnes attendues
- Les options de personnalisation
- Le développement de l'extension
- Les fonctionnalités avancées

## Raccourcis clavier
- **Ctrl+Enter** : Copier les données (depuis la zone de texte)

---

**Version** : 1.0.0
**Support** : Consultez le README.md pour plus d'informations
