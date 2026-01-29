# PdfToWord

# Description
PdfToWord est un petit script Python qui extrait les tableaux d’un fichier PDF et génère :
- un fichier CSV (export des tableaux Camelot)
- un fichier Word .docx contenant les tableaux

Le CSV et le DOCX sont enregistrés au même endroit que le PDF, avec le même nom.

# Fonctionnement
Le script fait deux étapes :
1. Lecture du PDF et extraction des tableaux avec Camelot
2. Création d’un document Word et insertion des tableaux extraits

# Structure du projet
- main.py : contient la classe PdfToWord et le point d’entrée du programme

# Prérequis
- Python 3
- Dépendances Python :
  - camelot-py
  - python-docx
  - ghostscript (souvent nécessaire selon l’installation de Camelot)
  - opencv-python (souvent nécessaire selon le mode utilisé)

# Installation
# 1) Créer et activer un environnement (optionnel mais recommandé)
Conda :
- conda create -n dev python=3.10
- conda activate dev

# 2) Installer les dépendances
Pip :
- pip install camelot-py[cv] python-docx

Si Camelot ne fonctionne pas, installe Ghostscript puis relance.

# Utilisation
# 1) Lancer le script
Dans le dossier où se trouve main.py :

- python main.py

# 2) Entrer le chemin du PDF
Quand le programme demande :
The path of your pdf:

Entre un chemin complet, par exemple :
C:\Users\Lucas\Desktop\test_5_tables.pdf

# Sorties générées
Si ton PDF s’appelle :
C:\...\mon_fichier.pdf

Le script génère :
- C:\...\mon_fichier.csv
- C:\...\mon_fichier.docx

Si compress=True, l’export CSV peut être généré sous forme de fichier zip selon Camelot.

# Exemple de code
Le script principal :
- demande le chemin du PDF
- exécute get_pdf_text()
- exécute text_to_word()

# Limitations
- Camelot fonctionne mieux avec des PDFs contenant des tableaux “vectoriels” (pas des scans images).
- Selon le type de PDF, tu peux devoir changer le mode Camelot :
  - flavor="lattice" pour les tableaux avec lignes visibles
  - flavor="stream" pour les tableaux sans lignes

# Dépannage
# Erreur: Cannot save file into a non-existent directory
Vérifie que tu donnes bien un fichier PDF (avec .pdf) et non un dossier.
