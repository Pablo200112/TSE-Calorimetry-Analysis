| Script | Purpose (EN) | Objectif (FR) |
| :--- | :--- | :--- |
| `zt_outlier_cleaner.py` | Detect/cap metabolic outliers (MAD/Z-Score) and format data into a 4-day continuous ZT chronology. | Détecte/plafonne les outliers (MAD/Z-Score) et formate les données en une chronologie ZT continue sur 4 jours. |
| `TSE_Add_EE.py` | Add Energy Expenditure column from VO2 and animal weights. | Ajoute la colonne Dépense Énergétique à partir de VO2 et des poids animaux. |
| `TSE_merge_excel.py` | Merge two Excel files by animal. | Fusionne deux fichiers Excel par animal. |
| `TSE_All-Graph_Raw.py` | Generate 15-min raw or smoothed graphs per animal. | Génère des graphiques 15-min bruts ou lissés par animal. |
| `TSE_All-Graph_mean.py` | Generate graphs using averaged data. | Génère des graphiques à partir de données moyennées. |
| `TSE_One_Day_mean.py` | Compute hourly averages/sums for a selected day. | Calcul des moyennes et sommes horaires pour un jour sélectionné. |
| `TSE_One_Day_raw.py` | Extract raw 15-min data for a selected day. | Extraction des données brutes 15-min pour un jour sélectionné. |
| `TSE_4_Days_raw.py` | Extract raw 15-min data for four consecutive days. | Extraction des données brutes 15-min pour quatre jours consécutifs. |

---

## 🔬 Core Features: Statistical Outlier Mitigation

The script `zt_outlier_cleaner.py` embeds a advanced statistical filter to clean food intake data (`Feed` column) before restructuring:
* **Robust Statistics (MAD):** Instead of using standard deviation—which is heavily biased by anomalous data points—it uses the **Median Absolute Deviation (MAD)**.
* **Modified Z-Score:** It applies a strict modified Z-score threshold of $3.5$. Any value exceeding this robust biological maximum is dynamically capped (ceiling effect), preserving the overall structure of the dataset without throwing away entire data rows.

---

## ⚙️ Execution Pipeline / Ordre d'Exécution

⚠️ **CRITICAL:** To process your data correctly, please follow this specific order / Pour que l'analyse fonctionne correctement, vous devez suivre cet ordre précis :

1. **Step 1 / Étape 1 : `TSE_merge_excel.py`**
   * *EN:* Merge your source files by animal first.
   * *FR:* Fusionnez d'abord vos fichiers sources par animal.

2. **Step 2 / Étape 2 : `TSE_Add_EE.py`**
   * *EN:* Run this script on the merged file to compute Energy Expenditure.
   * *FR:* Appliquez ce script sur le fichier fusionné pour calculer la dépense énergétique.

3. **Step 3 / Étape 3 : `zt_outlier_cleaner.py`**
   * *EN:* Clean French localization formatting, automatically detect and cap food consumption outliers using a robust Modified Z-Score (MAD method), and rebuild the continuous ZT framework (anchoring 07:00 AM as ZT00) tracking biological day switches per animal.
   * *FR:* Corrige les virgules de localisation, détecte et plafonne automatiquement les outliers de consommation via la méthode robuste du Z-Score Modifié (MAD), et reconstruit la matrice chronologique continue (07h00 = ZT00) en suivant les sauts de jours biologiques par animal.

4. **Step 4 / Étape 4 : Downstream Analysis & Graphing [Your Choice / Au choix]**
   * *EN:* Run any of the remaining scripts depending on the analysis or graph you need:
   * *FR:* Utilisez ensuite n'importe quel autre script restant selon votre besoin :
     * `TSE_All-Graph_Raw.py` : 15-min raw/smoothed graphs per animal.
     * `TSE_All-Graph_mean.py` : Graphs using averaged data.
     * `TSE_One_Day_mean.py` : Hourly averages/sums for a selected day.
     * `TSE_One_Day_raw.py` : Raw 15-min data for a selected day.
     * `TSE_4_Days_raw.py` : Raw 15-min data for 4 consecutive days.

---

## 📋 Requirements / Prérequis

* **Python 3.x**
* **Required Packages:** `pandas`, `numpy`, `matplotlib`, `openpyxl`, `tkinter`

To install all required packages at once, run / Pour installer tous les packages requis, exécutez :
```bash
pip install pandas numpy matplotlib openpyxl tk
