# Script de génération des listes de distributions amap

## Installation

Lancer le script suivant : 
  ```bash
  pip install -r requirements.txt
  ```

Installer les logiciels suivants : 
(1) Libre office (manipulation de fichiers xls)
(2) https://wkhtmltopdf.org/downloads.html (génération de pdfs)

Attention à bien vérifier que libre office est ajouté aux variables d'environnement.  

## Génération des fichiers de distribution
- **Etape 1**: Charger les fichiers xls dans le répertoire

- **Etape 2**: Ouvrir un terminal et convertir les fichiers xls en fichiers xlsx
  ```bash
  sh convertor.sh
  ```
- **Etape 3**: Depuis ce même terminal, lancer la génération du fichier xlsx et du pdf à imprimer pour la permanence de la semaine
  ```bash
  python src/main.py
  ```
