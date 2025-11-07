# -*- coding: utf-8 -*-
"""
4-Day Calorimetry Analysis
One graph per parameter, per animal
Cycles: Jour1 (LD12:12), Jour2 (Dark/Dark), Jour3 (LD1:1), Jour4 (LD12:12)
"""

import os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from tkinter import Tk, filedialog, simpledialog
from datetime import timedelta

# --------------------------
# 📂 Select Excel file
Tk().withdraw()
file_path = filedialog.askopenfilename(
    title="Select your merged Excel file",
    filetypes=[("Excel Files", "*.xlsx *.xls")]
)
if not file_path:
    raise FileNotFoundError("❌ No file selected.")

# --------------------------
# 📅 Choose starting day
start_day_str = simpledialog.askstring("Start Date", "Enter the START date (YYYY-MM-DD)")
start_day = pd.to_datetime(start_day_str).date()

# --------------------------
# ⚙️ Define the 4 cycles
cycles = [
    ("Jour1_LD12-12", "3"),
    ("Jour2_DarkDark", "2"),
    ("Jour3_LD1-1", "1"),
    ("Jour4_LD12-12", "3"),
]

# Output folder
output_root = r"C:\Users\pablo\OneDrive\Bureau\Program Output"
os.makedirs(output_root, exist_ok=True)
base_name = os.path.splitext(os.path.basename(file_path))[0]

all_days_data = []

# --------------------------
# Loop through the 4 days
for i, (cycle_name, cycle_code) in enumerate(cycles):
    day = start_day + timedelta(days=i)
    start_period = pd.to_datetime(str(day)) + pd.Timedelta(hours=7)
    end_period = start_period + pd.Timedelta(hours=24)

    # Read sheet
    df = pd.read_excel(file_path, sheet_name='PS 2025 02')
    df.columns = df.columns.str.strip()

    # ✅ Robust renaming (from your working script)
    df = df.rename(columns={
        "PS 2025 02": "Date",
        "Unnamed: 1": "Time",
        "TX002": "Animal",
        "Unnamed: 13": "RER",
        "Unnamed: 14": "XT_YT",
        "Unnamed: 15": "Feed"
    })

    # ✅ Force EE column if needed (column Q)
    if len(df.columns) >= 17:
        ee_col_name = df.columns[16]
        df = df.rename(columns={ee_col_name: "EE"})
    else:
        df["EE"] = None

    # ✅ Keep only relevant columns
    useful_columns = ["Date", "Time", "Animal", "RER", "XT_YT", "Feed", "EE"]
    df = df[[c for c in useful_columns if c in df.columns]].copy()

    # ✅ Ensure Animal column exists
    if "Animal" not in df.columns:
        raise KeyError("❌ 'Animal' column not found in Excel sheet.")

    # ✅ Filter and clean data
    df = df[pd.to_numeric(df["Animal"], errors="coerce").notna()]
    df["Animal"] = df["Animal"].astype(int)

    # ✅ Create DateTime safely
    if "Date" in df.columns and "Time" in df.columns:
        df["DateTime"] = pd.to_datetime(
            df["Date"].astype(str).str.strip() + " " + df["Time"].astype(str).str.strip(),
            errors="coerce"
        )
    else:
        raise KeyError("❌ 'Date' and/or 'Time' columns missing (check Excel headers).")

    # ✅ Convert numeric columns
    for col in ["RER", "XT_YT", "Feed", "EE"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    df = df.sort_values(["Animal", "DateTime"]).copy()

    # ✅ Feed_diff and normalization
    if "Feed" in df.columns:
        df["Feed_diff"] = df.groupby("Animal")["Feed"].diff().clip(lower=0)
    else:
        df["Feed_diff"] = pd.NA

    if "XT_YT" in df.columns:
        df["XT_YT"] = df["XT_YT"] / 8000

    # ✅ Select current day data
    df_day = df[(df["DateTime"] >= start_period) & (df["DateTime"] < end_period)].copy()

    # ✅ Aggregate
    agg_dict = {}
    if "RER" in df_day.columns: agg_dict["RER"] = "mean"
    if "XT_YT" in df_day.columns: agg_dict["XT_YT"] = "mean"
    if "Feed_diff" in df_day.columns: agg_dict["Feed_diff"] = "sum"
    if "EE" in df_day.columns: agg_dict["EE"] = "sum"

    if not agg_dict:
        print(f"⚠️ No aggregatable columns found for day {i+1}")
        continue

    df_pivot = df_day.groupby(["DateTime", "Animal"]).agg(agg_dict).reset_index()
    df_pivot["Cycle"] = cycle_name
    df_pivot["CycleType"] = cycle_code
    all_days_data.append(df_pivot)

# --------------------------
# Combine all days
df_all = pd.concat(all_days_data, ignore_index=True)
animals = sorted(df_all["Animal"].unique())

# Colors for parameters
param_colors = {"RER": "blue", "XT_YT": "red", "Feed_diff": "green", "EE": "purple"}

# --------------------------
# Function to shade light/dark
def shade_light_cycle(ax, start_time, cycle_type):
    if cycle_type == "1":  # LD1:1
        for h in range(0, 24, 2):
            ax.axvspan(start_time + pd.Timedelta(hours=h + 1),
                       start_time + pd.Timedelta(hours=h + 2),
                       color='gray', alpha=0.15)
    elif cycle_type == "2":  # DarkDark
        ax.axvspan(start_time, start_time + pd.Timedelta(hours=24),
                   color='gray', alpha=0.25)
    elif cycle_type == "3":  # LD12:12
        ax.axvspan(start_time + pd.Timedelta(hours=12),
                   start_time + pd.Timedelta(hours=24),
                   color='gray', alpha=0.25)

# --------------------------
# Loop through all animals and parameters
for animal in animals:
    df_animal = df_all[df_all["Animal"] == animal]

    for param, color in param_colors.items():
        if param not in df_animal.columns or df_animal[param].isna().all():
            print(f"⚠️ Column '{param}' missing or empty for Animal {animal}, skipping.")
            continue

        fig, ax = plt.subplots(figsize=(16,6))
        ax.plot(df_animal["DateTime"], df_animal[param],
                color=color, linewidth=2, label=param)

        for i, (cycle_name, cycle_code) in enumerate(cycles):
            day_start = pd.to_datetime(str(start_day + timedelta(days=i))) + pd.Timedelta(hours=7)
            shade_light_cycle(ax, day_start, cycle_code)

        ax.set_title(f"Animal {animal} - {param} over 4 Days (7h→7h)", fontsize=14)
        ax.set_xlabel("Date and Time")
        ax.set_ylabel(param)
        ax.xaxis.set_major_locator(mdates.HourLocator(interval=6))
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%m-%d %Hh'))
        ax.grid(True, linestyle='--', alpha=0.6)
        plt.xticks(rotation=45)
        plt.tight_layout()

        save_name = f"Animal{animal}_{param}_4Days.png"
        plt.savefig(os.path.join(output_root, save_name))
        plt.close()

print(f"\n✅ All graphs generated for {len(animals)} animals in: {output_root}")
