FlowBoard - Visual Command Center
==================================

Un dashboard élégant pour organiser et lancer vos raccourcis (fichiers, dossiers, URLs) en un clic.

FONCTIONNALITES
---------------
- Spaces : Organisez vos raccourcis dans des catégories personnalisées
- Recherche double mode : Google ou filtrage interne des raccourcis
- Favicons auto : Icônes de sites web téléchargées automatiquement
- Icônes natives : Support des .lnk Windows et icônes système
- Notes contextuelles : Infobulles au survol des cartes
- UI moderne : Thème sombre, effets glassmorphism, animations fluides
- Drag & Drop : Réorganisez les cartes ou glissez-déposez depuis l'explorateur
- Undo : Annulez vos dernières actions avec Ctrl+Z
- Raccourcis clavier : Ctrl+F, Ctrl+N, Escape, Ctrl+Z
- Persistance JSON : Sauvegarde automatique de vos données

INSTALLATION
------------
Prérequis : Python 3.9+

1. Cloner le projet :
   git clone https://github.com/HocineBenbara/FlowBoard.git
   cd FlowBoard

2. Créer l'environnement virtuel :
   python -m venv venv
   # Windows :
   venv\Scripts\activate
   # macOS/Linux :
   source venv/bin/activate

3. Installer les dépendances :
   pip install -r requirements.txt

   [Windows uniquement] Pour le support des raccourcis .lnk :
   pip install pywin32

LANCEMENT
---------
python main.py

UTILISATION
-----------
Ajouter un raccourci :
- Cliquez sur "+ Add shortcut" ou faites Ctrl+N
- Collez une URL ou un chemin de fichier/dossier
- Optionnel : nom personnalisé + note contextuelle
- Validez → la carte apparaît dans la grille

Organiser :
- Glissez-déposez les cartes pour les réordonner
- Clic droit sur une carte : Ouvrir / Modifier / Déplacer / Supprimer
- Glissez une carte sur un Space dans la sidebar pour la catégoriser

Rechercher :
- Mode Google (par défaut) : tapez et Entrée → ouvre Google
- Mode FlowBoard : basculez sur "FlowBoard" pour filtrer vos raccourcis

Raccourcis clavier :
- Ctrl+F : Focus sur la barre de recherche
- Ctrl+N : Nouveau raccourci
- Ctrl+Z : Annuler la dernière action
- Escape : Effacer la recherche / Fermer un dialog

STRUCTURE DU PROJET
-------------------
FlowBoard/
├── main.py              # Point d'entrée + classe principale
├── requirements.txt     # Dépendances Python
├── icon.ico             # Icône de l'application (Windows)
├── QATracker.spec       # Spec PyInstaller (optionnel)
├── .gitignore           # Fichiers exclus du versionning
├── PySide6_data/        # Dossier généré à l'exécution
│   ├── shortcuts.json   # Base de données des raccourcis
│   ├── categories.json  # Configuration des Spaces
│   └── favicons/        # Cache des icônes web
└── venv/                # Environnement virtuel (ignoré par Git)

CONFIGURATION AVANCÉE
---------------------
Personnaliser le thème :
- Modifiez le dictionnaire C dans la classe FlowBoard :
  C = {
      "primary": "#6c5ce7",      # Couleur principale
      "primary_dark": "#3d2d9c", # Variante sombre
      # ... autres couleurs
  }

Ajuster la grille :
- CARD_SIZE = 122 : taille des cartes en pixels
- CARD_STRIDE = 134 : espacement entre cartes
- COLS_DEFAULT = 5 : nombre de colonnes par défaut

DÉVELOPPEMENT
-------------
Lancer en mode debug :
  python -u main.py

Compiler en exécutable (Windows) :
  pip install pyinstaller
  pyinstaller QATracker.spec
  # L'exécutable sera dans dist/

CONTRIBUTER
-----------
1. Fork le projet
2. Créez votre branche : git checkout -b feature/ma-fonctionnalite
3. Committez : git commit -m 'Ajout: ma fonctionnalité'
4. Poussez : git push origin feature/ma-fonctionnalite
5. Ouvrez une Pull Request

LICENCE
-------
Ce projet est distribué sous licence MIT.

SUPPORT
-------
- Bug : https://github.com/HocineBenbara/FlowBoard/issues
- Idées : https://github.com/HocineBenbara/FlowBoard/discussions
- Contact : hocine.benbara@proton.me

---
Créé par Hocine Benbara - https://github.com/HocineBenbara
