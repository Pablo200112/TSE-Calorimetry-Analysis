# -*- coding: utf-8 -*-
"""
4-Day Calorimetry Analysis
Hourly aggregation (LD11-compatible)

Author: Pablo SAIDI
"""

import os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from tkinter import Tk, filedialog, simpledialog, messagebox
from datetime import timedelta

# --------------------------
# 📂 Select Excel file
# --------------------------
Tk().withdraw()
file_path = filedialog.askopenfilename(
    title="Select your merged Excel file",
    filetypes=[("Excel Files", "*.xlsx *.xls")]
)
if not file_path:
    raise FileNotFoundError("❌ No file selected.")

# --------------------------
# 📅 Choose starting day
# --------------------------
start_day_str = simpledialog.askstring(
    "Start Date",
    "Enter the START date (YYYY-MM-DD)\n(analysis starts at 7:00)"
)
start_day = pd.to_datetime(start_day_str).date()

# --------------------------
# ⏱️ Timestamp alignment
# --------------------------
timestamp_mode = simpledialog.askstring(
    "Timestamp Position",
    "Sampling window = 15 min\n\n"
    "1 = beginning of window\n"
    "2 = center of window (recommended)\n"
    "3 = end of window\n\n"
    "Enter 1, 2 or 3:"
)

if timestamp_mode not in ["1", "2", "3"]:
    raise ValueError("❌ Invalid choice.")

if timestamp_mode == "1":
    timestamp_shift = pd.Timedelta(minutes=15)
elif timestamp_mode == "2":
    timestamp_shift = pd.Timedelta(minutes=7, seconds=30)
else:
    timestamp_shift = pd.Timedelta(0)

# --------------------------
# ⚙️ Define 4 cycles
# --------------------------
cycles = [
    ("Day1_LD12-12", "3"),
    ("Day2_DD", "2"),
    ("Day3_LD1-1", "1"),
    ("Day4_LD12-12", "3"),
]

# --------------------------
# 📁 Output folder
# --------------------------
output_root = r"D:\pablo.SAIDI\Desktop\Sortie programme calo"
os.makedirs(output_root, exist_ok=True)
base_name = os.path.splitext(os.path.basename(file_path))[0]

# --------------------------
# ❓ Feed filtering option
# --------------------------
root = Tk()
root.withdraw()
exclude_feed = messagebox.askyesno(
    "Feed filtering",
    "Exclude Feed_diff values > 2 g ?"
)
root.destroy()

# --------------------------
# 📊 Read Excel
# --------------------------
df = pd.read_excel(file_path, sheet_name='PS 2025 02')
df.columns = df.columns.str.strip()

df = df.rename(columns={
    "PS 2025 02": "Date",
    "Unnamed: 1": "Time",
    "TX002": "Animal",
    "Unnamed: 13": "RER",
    "Unnamed: 14": "XT_YT",
    "Unnamed: 15": "Feed"
})

# EE
if len(df.columns) >= 17:
    df = df.rename(columns={df.columns[16]: "EE"})
else:
    df["EE"] = None

# --------------------------
# 🧹 Cleaning
# --------------------------
df = df[pd.to_numeric(df["Animal"], errors="coerce").notna()]
df["Animal"] = df["Animal"].astype(int)

df["DateTime"] = pd.to_datetime(
    df["Date"].astype(str).str.strip() + " " +
    df["Time"].astype(str).str.strip(),
    errors="coerce"
)

# Shift timestamps
df["DateTime"] = df["DateTime"] - timestamp_shift

for col in ["RER", "XT_YT", "Feed", "EE"]:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

df = df.sort_values(["Animal", "DateTime"]).copy()

# --------------------------
# 🍽️ Feed diff + filtering
# --------------------------
df["Feed_diff"] = df.groupby("Animal")["Feed"].diff()
df.loc[df["Feed_diff"] < 0, "Feed_diff"] = 0

if exclude_feed:
    df["Feed_diff"] = df["Feed_diff"].where(df["Feed_diff"] <= 2, None)

# --------------------------
# Normalize XT+YT
# --------------------------
df["XT_YT"] = df["XT_YT"] / 8000

# --------------------------
# 🌗 Light cycle shading
# --------------------------
def shade_light_cycle(ax, start_time, cycle_type):
    if cycle_type == "1":  # LD1:1
        for h in range(0, 24, 2):
            ax.axvspan(start_time + pd.Timedelta(hours=h+1),
                       start_time + pd.Timedelta(hours=h+2),
                       color='gray', alpha=0.15)
    elif cycle_type == "2":  # DD
        ax.axvspan(start_time, start_time + pd.Timedelta(hours=24),
                   color='gray', alpha=0.25)
    elif cycle_type == "3":  # LD12:12
        ax.axvspan(start_time + pd.Timedelta(hours=12),
                   start_time + pd.Timedelta(hours=24),
                   color='gray', alpha=0.25)

# --------------------------
# 🔁 Loop over days
# --------------------------
all_days = []

for i, (cycle_name, cycle_code) in enumerate(cycles):

    day = start_day + timedelta(days=i)
    start_period = pd.to_datetime(str(day)) + pd.Timedelta(hours=7)
    end_period = start_period + pd.Timedelta(hours=24)

    df_day = df[
        (df["DateTime"] >= start_period) &
        (df["DateTime"] < end_period)
    ].copy()

    # Relative hour
    df_day["Relative_Hour"] = (
        (df_day["DateTime"] - start_period)
        .dt.total_seconds() // 3600
    ).astype(int)

    # Hourly aggregation
    agg = {}
    if "RER" in df_day.columns: agg["RER"] = "mean"
    if "XT_YT" in df_day.columns: agg["XT_YT"] = "sum"
    if "Feed_diff" in df_day.columns: agg["Feed_diff"] = "sum"
    if "EE" in df_day.columns: agg["EE"] = "sum"

    df_hour = (
        df_day
        .groupby(["Relative_Hour", "Animal"])
        .agg(agg)
        .reset_index()
    )

    df_hour["DateTime"] = (
        start_period
        + pd.to_timedelta(df_hour["Relative_Hour"], unit="h")
        + pd.to_timedelta(0.5, unit="h")
    )

    df_hour["Cycle"] = cycle_name
    df_hour["CycleType"] = cycle_code

    all_days.append(df_hour)

# --------------------------
# 📦 Combine all days
# --------------------------
df_all = pd.concat(all_days, ignore_index=True)
animals = sorted(df_all["Animal"].unique())

# --------------------------
# 📈 Plot
# --------------------------
param_colors = {
    "RER": "blue",
    "XT_YT": "red",
    "Feed_diff": "green",
    "EE": "purple"
}

for animal in animals:
    df_a = df_all[df_all["Animal"] == animal]

    for param, color in param_colors.items():
        if param not in df_a.columns or df_a[param].isna().all():
            continue

        fig, ax = plt.subplots(figsize=(16, 6))
        ax.plot(df_a["DateTime"], df_a[param],
                color=color, linewidth=2)

        for i, (_, cycle_code) in enumerate(cycles):
            day_start = (
                pd.to_datetime(str(start_day + timedelta(days=i)))
                + pd.Timedelta(hours=7)
            )
            shade_light_cycle(ax, day_start, cycle_code)

        ax.set_title(f"Animal {animal} – {param} (4 days, hourly)")
        ax.set_xlabel("DateTime")
        ax.set_ylabel(param)
        ax.grid(True, linestyle="--", alpha=0.6)
        ax.xaxis.set_major_locator(mdates.HourLocator(interval=6))
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%m-%d %Hh'))
        plt.xticks(rotation=45)
        plt.tight_layout()

        save_name = f"Animal{animal}_{param}_4Days_hourly.png"
        plt.savefig(os.path.join(output_root, save_name))
        plt.close()

print("\n✅ All hourly 4-day graphs generated successfully.")
