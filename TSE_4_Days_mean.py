# -*- coding: utf-8 -*-
"""
4-Day Calorimetry Analysis
Hourly aggregation (LD11-compatible)
Timestamp correction + Feed filtering
Y-axis scaling modes: auto / global / manual
Created by Pablo SAIDI
"""

import os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from tkinter import Tk, filedialog, simpledialog, messagebox
from datetime import timedelta

# ======================================================
# 📂 Select Excel file
# ======================================================
Tk().withdraw()
file_path = filedialog.askopenfilename(
    title="Select your merged Excel file",
    filetypes=[("Excel Files", "*.xlsx *.xls")]
)
if not file_path:
    raise FileNotFoundError("❌ No file selected.")

# ======================================================
# 📅 Choose starting day
# ======================================================
start_day_str = simpledialog.askstring(
    "Start date",
    "Enter START date (YYYY-MM-DD)\nAnalysis runs 7 AM → 7 AM for 4 days"
)
start_day = pd.to_datetime(start_day_str).date()

# ======================================================
# ⏱️ Timestamp alignment
# ======================================================
timestamp_mode = simpledialog.askstring(
    "Timestamp position",
    "Sampling window = 15 min\n\n"
    "1 = beginning of window\n"
    "2 = center of window (recommended)\n"
    "3 = end of window\n\n"
    "Enter 1, 2 or 3"
)

if timestamp_mode == "1":
    timestamp_shift = pd.Timedelta(minutes=15)
elif timestamp_mode == "2":
    timestamp_shift = pd.Timedelta(minutes=7, seconds=30)
elif timestamp_mode == "3":
    timestamp_shift = pd.Timedelta(0)
else:
    raise ValueError("Invalid timestamp choice.")

# ======================================================
# 🧪 Feed filtering
# ======================================================
root = Tk()
root.withdraw()
filter_feed = messagebox.askyesno(
    "Feed filtering",
    "Exclude Feed_diff values > 2 g ?"
)
root.destroy()

# ======================================================
# 📏 Y-axis scaling mode
# ======================================================
root = Tk()
root.withdraw()
y_scale_mode = simpledialog.askstring(
    "Y-axis scaling",
    "Choose Y-axis scaling mode:\n\n"
    "1 = Autoscale (per animal)\n"
    "2 = Same scale (auto, all animals)\n"
    "3 = Manual scale (user-defined)\n\n"
    "Enter 1, 2 or 3"
)
root.destroy()

if y_scale_mode not in ["1", "2", "3"]:
    raise ValueError("Invalid Y-axis scaling choice.")

# ======================================================
# ⚙️ Experimental cycles
# ======================================================
cycles = [
    ("Day1_LD12-12", "3"),
    ("Day2_DD", "2"),
    ("Day3_LD1-1", "1"),
    ("Day4_LD12-12", "3"),
]

# ======================================================
# 📁 Output folder
# ======================================================
output_root = r"D:\pablo.SAIDI\Desktop\Sortie programme calo"
os.makedirs(output_root, exist_ok=True)
base_name = os.path.splitext(os.path.basename(file_path))[0]

all_days_data = []

# ======================================================
# 🔁 Loop over 4 days
# ======================================================
for i, (cycle_name, cycle_code) in enumerate(cycles):

    day = start_day + timedelta(days=i)
    start_period = pd.to_datetime(str(day)) + pd.Timedelta(hours=7)
    end_period = start_period + pd.Timedelta(hours=24)

    # --------------------------
    # 📊 Read Excel
    df = pd.read_excel(file_path, sheet_name='2em PS 2025 01')
    df.columns = df.columns.str.strip()

    df = df.rename(columns={
        df.columns[0]: "Date",
        df.columns[1]: "Time",
        "TX002": "Animal",
        "Unnamed: 13": "RER",
        "Unnamed: 14": "XT_YT",
        "Unnamed: 15": "Feed"
    })

    if len(df.columns) >= 17:
        df = df.rename(columns={df.columns[16]: "EE"})
    else:
        df["EE"] = pd.NA

    df = df[pd.to_numeric(df["Animal"], errors="coerce").notna()]
    df["Animal"] = df["Animal"].astype(int)

    df["DateTime"] = pd.to_datetime(
        df["Date"].astype(str) + " " + df["Time"].astype(str),
        errors="coerce"
    )

    # ⏱️ Timestamp correction
    df["DateTime"] = df["DateTime"] - timestamp_shift

    # Numeric conversion
    for col in ["RER", "XT_YT", "Feed", "EE"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df = df.sort_values(["Animal", "DateTime"]).copy()

    # --------------------------
    # 🍽️ Feed diff
    df["Feed_diff"] = df.groupby("Animal")["Feed"].diff()
    df.loc[df["Feed_diff"] < 0, "Feed_diff"] = 0

    if filter_feed:
        df.loc[df["Feed_diff"] > 2, "Feed_diff"] = pd.NA

    # Normalize activity
    df["XT_YT"] = df["XT_YT"] / 8000

    # --------------------------
    # ⏱️ Select day window
    df_day = df[(df["DateTime"] >= start_period) &
                (df["DateTime"] < end_period)].copy()

    # Relative hour
    df_day["Relative_Hour"] = (
        (df_day["DateTime"] - start_period)
        .dt.total_seconds() // 3600
    ).astype(int)

    # --------------------------
    # 🧮 Hourly aggregation
    agg = {
        "RER": "mean",
        "XT_YT": "sum",
        "Feed_diff": "sum",
        "EE": "sum"
    }

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

    all_days_data.append(df_hour)

# ======================================================
# 🔗 Combine all days
# ======================================================
df_all = pd.concat(all_days_data, ignore_index=True)
animals = sorted(df_all["Animal"].unique())

# ======================================================
# 📏 Y-axis limits
# ======================================================
global_y_limits = {}
manual_y_limits = {}

if y_scale_mode == "2":
    for param in ["RER", "XT_YT", "Feed_diff", "EE"]:
        ymin = df_all[param].min(skipna=True)
        ymax = df_all[param].max(skipna=True)
        margin = 0.05 * (ymax - ymin) if ymax > ymin else 0
        global_y_limits[param] = (ymin - margin, ymax + margin)

elif y_scale_mode == "3":
    for param in ["RER", "XT_YT", "Feed_diff", "EE"]:
        ymin = simpledialog.askfloat(
            f"{param} Y-min",
            f"Enter Y-axis MIN for {param} (Cancel = autoscale)"
        )
        ymax = simpledialog.askfloat(
            f"{param} Y-max",
            f"Enter Y-axis MAX for {param} (Cancel = autoscale)"
        )
        if ymin is not None and ymax is not None:
            manual_y_limits[param] = (ymin, ymax)

# ======================================================
# 🌗 Light cycle shading
# ======================================================
def shade_light_cycle(ax, start, cycle):
    if cycle == "1":  # LD 1:1
        for h in range(0, 24, 2):
            ax.axvspan(start + pd.Timedelta(hours=h+1),
                       start + pd.Timedelta(hours=h+2),
                       color='gray', alpha=0.2)
    elif cycle == "2":  # DD
        ax.axvspan(start, start + pd.Timedelta(hours=24),
                   color='gray', alpha=0.3)
    elif cycle == "3":  # LD 12:12
        ax.axvspan(start + pd.Timedelta(hours=12),
                   start + pd.Timedelta(hours=24),
                   color='gray', alpha=0.3)

# ======================================================
# 📈 Plot
# ======================================================
param_colors = {
    "RER": "blue",
    "XT_YT": "red",
    "Feed_diff": "green",
    "EE": "purple"
}

for animal in animals:
    df_a = df_all[df_all["Animal"] == animal]

    for param, color in param_colors.items():
        if df_a[param].isna().all():
            continue

        fig, ax = plt.subplots(figsize=(16, 6))
        ax.plot(df_a["DateTime"], df_a[param], color=color, linewidth=2)

        for i, (_, cycle_code) in enumerate(cycles):
            day_start = pd.to_datetime(str(start_day + timedelta(days=i))) + pd.Timedelta(hours=7)
            shade_light_cycle(ax, day_start, cycle_code)

        # Apply Y-scale
        if y_scale_mode == "2" and param in global_y_limits:
            ax.set_ylim(global_y_limits[param])
        elif y_scale_mode == "3" and param in manual_y_limits:
            ax.set_ylim(manual_y_limits[param])

        ax.set_title(f"Animal {animal} – {param} (4 days)")
        ax.set_xlabel("Time")
        ax.set_ylabel(param)
        ax.grid(True, linestyle="--", alpha=0.6)
        ax.xaxis.set_major_locator(mdates.HourLocator(interval=6))
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%m-%d %Hh'))
        plt.xticks(rotation=45)
        plt.tight_layout()

        plt.savefig(os.path.join(
            output_root,
            f"Animal{animal}_{param}_4days_hourly.png"
        ))
        plt.close()

print("\n✅ All figures generated successfully.")
