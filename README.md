TSE Calorimetry Analysis Scripts
🇬🇧 English Version
Overview

This repository contains a set of Python scripts designed for calorimetry data analysis of animals. The scripts process raw Excel data, calculate energy expenditure, merge files, and generate detailed plots of RER, activity, feeding, and energy expenditure. All scripts use Python 3 with pandas, matplotlib, openpyxl, and tkinter for GUI file selection.

Scripts

TSE_Add_EE.py

Adds an Energy Expenditure (EE) column to your Excel file.

Reads VO2 values and animal weights automatically and calculates EE in kcal/h.

Output: Excel file with a new column Energy expenditure [kcal/h].

TSE_merge_excel.py

Merges two Excel files containing animal data.

Inserts data from the second file right after each corresponding animal in the first file.

Output: Final merged Excel file.

TSE_All-Graph_Raw.py

Generates 15-min raw or smoothed graphs for all animals.

Optionally applies a 1-hour rolling mean smoothing.

Outputs: individual and global graphs per metric (RER, XT+YT, Feed, EE) and a processed Excel file.

TSE_One_Day_mean.py

Computes hourly averages and sums for a selected 7 AM → 7 AM period.

Plots multi-axis and individual metric graphs for each animal and global plots for all animals.

Supports different light cycle types: LD1:1, DD, LD12:12.

Output: Excel file with hourly data and graphs.

TSE_One_Day_raw.py

Extracts raw 15-min data for a selected 7 AM → 7 AM period without averaging.

Generates multi-axis and individual metric plots for each animal.

Supports light cycles as above.

Output: Excel file with raw data and graphs.

Requirements

Python 3.x

Packages:
