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
