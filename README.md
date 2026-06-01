| Script                  | Purpose (EN)                                               | Objectif (FR)                                                               |
| ----------------------- | ---------------------------------------------------------- | --------------------------------------------------------------------------- |
| `TSE_Add_EE.py`         | Add Energy Expenditure column from VO2 and animal weights. | Ajoute la colonne Dépense Énergétique à partir de VO2 et des poids animaux. |
| `TSE_merge_excel.py`    | Merge two Excel files by animal.                           | Fusionne deux fichiers Excel par animal.                                    |
| `TSE_All-Graph_Raw.py`  | Generate 15-min raw or smoothed graphs per animal.         | Génère des graphiques 15-min bruts ou lissés par animal.                    |
| `TSE_All-Graph_mean.py` | Generate graphs using averaged data.                       | Génère des graphiques à partir de données moyennées.                        |
| `TSE_One_Day_mean.py`   | Compute hourly averages/sums for a selected day.           | Calcul des moyennes et sommes horaires pour un jour sélectionné.            |
| `TSE_One_Day_raw.py`    | Extract raw 15-min data for a selected day.                | Extraction des données brutes 15-min pour un jour sélectionné.              |
| `TSE_4_Days_raw.py`     | Extract raw 15-min data for four consecutive days.         | Extraction des données brutes 15-min pour quatre jours consécutifs.         |


-Requirements / Prérequis:
Python 3.x
Packages: pandas, matplotlib, openpyxl, tkinter
Install required packages: pip install pandas matplotlib openpyxl tk


-Usage / Utilisation :
Open your IDE (Spyder recommended). / Ouvrir l’IDE (Spyder recommandé).
Run the script and follow the prompts for file selection and options.
Output Excel files and graphs will be saved automatically in the specified output folder.


-Output / Sortie :
Excel files with processed/calculated data.
Graphs for each animal (multi-axis, raw or averaged).
Global graphs comparing all animals.

"""
===============================================================================
                       TSE DATA PROCESSING PIPELINE
===============================================================================

ORDER OF EXECUTION / ORDRE D'EXÉCUTION :
----------------------------------------
⚠️  Pour que l'analyse fonctionne correctement, vous devez suivre cet ordre précis :
    To process your data correctly, please follow this specific order:

1️⃣  Step 1 / Étape 1 : TSE_merge_excel.py
    -> Fusionnez d'abord vos fichiers sources par animal.
    -> Merge your source files by animal first.

2️⃣  Step 2 / Étape 2 : TSE_Add_EE.py
    -> Appliquez ce script sur le fichier fusionné pour calculer la dépense énergétique.
    -> Run this script on the merged file to compute Energy Expenditure.

3️⃣  Step 3 / Étape 3 : [Your Choice / Au choix]
    -> Utilisez ensuite n'importe quel autre script restant selon votre besoin :
       Run any of the remaining scripts depending on the analysis or graph you need:
       - TSE_All-Graph_Raw.py  : 15-min raw/smoothed graphs per animal
       - TSE_All-Graph_mean.py : Graphs using averaged data
       - TSE_One_Day_mean.py   : Hourly averages/sums for a selected day
       - TSE_One_Day_raw.py    : Raw 15-min data for a selected day
       - TSE_4_Days_raw.py     : Raw 15-min data for 4 consecutive days

-------------------------------------------------------------------------------
REQUIREMENTS / PRÉREQUIS :
    pip install pandas matplotlib openpyxl tk
===============================================================================
"""

# --- Début de ton code existant (ex: import pandas as pd, etc.) ---
import pandas as pd
# ... rest of your code ...
