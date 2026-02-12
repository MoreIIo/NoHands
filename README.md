# NoHands - Extension Chrome

Extension Chrome qui permet de copier des données Excel (25 colonnes) et de les formater pour un collage facile dans des formulaires web.

## 📋 Description

NoHands simplifie le transfert de données propriétaire depuis Excel vers des formulaires web. L'extension reformate automatiquement les 25 colonnes de données pour faciliter le remplissage de formulaires.

## ✨ Fonctionnalités

- **Interface intuitive** : Popup accessible depuis l'icône de l'extension
- **Parsing automatique** : Détecte et valide les 25 colonnes attendues
- **Formatage optimisé** : Transforme les données pour un collage facile dans les formulaires
- **Bouton stylisé** : Design moderne avec icône clipboard et animations
- **Validation robuste** : Messages d'erreur clairs en cas de données invalides
- **Raccourci clavier** : Ctrl+Enter pour copier rapidement

## 📦 Installation

### Mode Développeur (Local)

1. **Cloner ou télécharger** ce projet
2. **Créer les icônes** (voir section Icônes ci-dessous)
3. Ouvrir Chrome et naviguer vers `chrome://extensions/`
4. Activer le **"Mode développeur"** (toggle en haut à droite)
5. Cliquer sur **"Charger l'extension non empaquetée"**
6. Sélectionner le dossier `NoHands`
7. L'extension apparaît dans la barre d'outils Chrome

## 🎨 Icônes

L'extension nécessite 4 icônes PNG dans le dossier `icons/` :
- `icon16.png` (16x16px)
- `icon32.png` (32x32px)
- `icon48.png` (48x48px)
- `icon128.png` (128x128px)

### Créer vos icônes

**Option 1 : Outils en ligne (Recommandé)**
- [Favicon.io](https://favicon.io/) - Générateur d'icônes gratuit
- [Canva](https://www.canva.com/) - Outil de design en ligne
- [GIMP](https://www.gimp.org/) - Logiciel gratuit et open-source

**Option 2 : Icônes temporaires**
Vous pouvez créer des carrés de couleur simple pour tester l'extension :
```bash
# Avec ImageMagick (si installé)
magick -size 16x16 xc:#667eea icons/icon16.png
magick -size 32x32 xc:#667eea icons/icon32.png
magick -size 48x48 xc:#667eea icons/icon48.png
magick -size 128x128 xc:#667eea icons/icon128.png
```

**Concept suggéré** : Clipboard avec grille Excel, ou main avec flèche de transfert

## 🚀 Utilisation

### 1. Copier les données depuis Excel
Dans votre fichier Excel, sélectionnez la ligne avec les 25 colonnes :
```
N° PROP (TW)  CIVILITE  NOM  PRENOM  ADRESSE LIGNE 1  ...
```

Copiez la ligne entière (Ctrl+C)

### 2. Ouvrir l'extension
Cliquez sur l'icône NoHands dans la barre d'outils Chrome

### 3. Coller les données
Collez vos données dans la zone de texte (Ctrl+V)

### 4. Copier le format transformé
- Cliquez sur le bouton **"COPIER"**, ou
- Utilisez le raccourci **Ctrl+Enter**

### 5. Coller dans le formulaire
Collez les données formatées dans votre formulaire web (Ctrl+V)

## 📊 Format des données

### Colonnes attendues (25 au total)
1. N° PROP (TW)
2. CIVILITE PROP
3. NOM PROP
4. PRENOM PROP
5. ADRESSE LIGNE 1 PROP
6. ADRESSE LIGNE 2 PROP
7. CP PROP
8. VILLE PROP
9. TELEPHONE DOMICILE PROP
10. TELEPHONE BUREAU PROP
11. TELEPHONE PORTABLE PROP
12. EMAIL PROP
13. IBAN PROP
14. FREQUENCE REGLT ACOMPTE PROP
15. FREQUENCE REEDITION PROP
16. MODE REGLT AU PROP
17. TAUX HONOS PROP
18. ASSURANCE GL (O/N)
19. TAUX ASSURANCE GLI
20. TAUX HONOS/ASSURANCE BASE 1
21. DECLARATION REVENUS FONCIERS ADRF (O/N)
22. TYPE GARANTIE
23. DATE DEBUT MANDAT PROP
24. NOM GESTIONNAIRE
25. PRENOM GESTIONNAIRE

### Exemple d'entrée (Excel)
```
12345	M.	DUPONT	Jean	12 Rue de la Paix		75001	Paris	0123456789	...
```

### Exemple de sortie (Formaté)
```
N° PROP (TW): 12345
CIVILITE PROP: M.
NOM PROP: DUPONT
PRENOM PROP: Jean
ADRESSE LIGNE 1 PROP: 12 Rue de la Paix
ADRESSE LIGNE 2 PROP:
CP PROP: 75001
VILLE PROP: Paris
...
```

## 🐛 Dépannage

### L'extension ne se charge pas
- Vérifiez que toutes les icônes sont présentes dans le dossier `icons/`
- Vérifiez les erreurs dans `chrome://extensions/`
- Assurez-vous que le fichier `manifest.json` est valide

### Le bouton reste désactivé
- La zone de texte doit contenir du texte pour activer le bouton
- Vérifiez que vous avez bien collé les données

### Erreur "Format invalide"
- Vérifiez que vous avez exactement 25 colonnes
- Les colonnes doivent être séparées par des tabulations (depuis Excel)
- Vérifiez qu'il n'y a pas de saut de ligne dans les données

### La copie ne fonctionne pas
- Vérifiez les permissions du navigateur
- Essayez de fermer et rouvrir le popup
- Vérifiez la console du navigateur (F12 sur le popup)

## 🛠️ Développement

### Structure du projet
```
NoHands/
├── manifest.json          # Configuration Chrome Manifest V3
├── popup.html            # Interface utilisateur
├── popup.js              # Logique métier
├── styles.css            # Design et animations
├── icons/                # Icônes de l'extension
│   ├── icon16.png
│   ├── icon32.png
│   ├── icon48.png
│   └── icon128.png
└── README.md             # Ce fichier
```

### Technologies
- **Vanilla JavaScript** (pas de framework)
- **Chrome Manifest V3** (standard actuel)
- **Clipboard API** moderne (async/await)
- **CSS moderne** (Flexbox, animations)

### Debugging
1. Ouvrir le popup de l'extension
2. Clic droit → **Inspecter**
3. DevTools s'ouvre avec la console pour le popup
4. Vérifier les logs et erreurs JavaScript

### Tests
Testez avec différents types de données :
- Ligne complète (25 colonnes remplies)
- Ligne avec colonnes vides
- Caractères spéciaux (é, è, ê, etc.)
- Valeurs très longues
- Mauvais nombre de colonnes (erreur attendue)

## 🔒 Sécurité et confidentialité

- ✅ **Aucune connexion réseau** : L'extension fonctionne entièrement en local
- ✅ **Aucun stockage de données** : Rien n'est sauvegardé
- ✅ **Aucun tracking** : Pas d'analytics ou de télémétrie
- ✅ **Permissions minimales** : Seulement `clipboardWrite`
- ✅ **Open source** : Code totalement transparent

## 📝 Licence

Ce projet est sous licence MIT. Libre d'utilisation, modification et distribution.

## 🤝 Contribution

Les contributions sont les bienvenues ! N'hésitez pas à :
- Signaler des bugs
- Proposer des améliorations
- Soumettre des pull requests

## 📧 Support

Pour toute question ou problème :
1. Vérifiez la section [Dépannage](#-dépannage)
2. Consultez les issues GitHub (si applicable)
3. Créez une nouvelle issue avec une description détaillée

---

**Version** : 1.0.0
**Dernière mise à jour** : Février 2026
