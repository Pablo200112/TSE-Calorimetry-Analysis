# -*- coding: utf-8 -*-
"""
Complete Script: Calorimetry Analysis LD11 (No hourly averaging)
Created by Pablo SAIDI
"""

import os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from tkinter import Tk, filedialog, simpledialog

# --------------------------
# 📂 Select Excel file
Tk().withdraw()
file_path = filedialog.askopenfilename(
    title="Select your merged Excel file",
    filetypes=[("Excel Files", "*.xlsx *.xls")]
)

if not file_path:
    raise FileNotFoundError("❌ No file selected. Please restart the script and choose an Excel file.")

print(f"✅ Selected file: {file_path}")

# --------------------------
# 🗓️ Choose start day (7 AM → 7 AM next day)
start_day_str = simpledialog.askstring(
    "Select Day",
    "Enter the START date of the period (YYYY-MM-DD)\n"
    "Example: 2025-10-15 to analyze from Oct 15th 7 AM to Oct 16th 7 AM"
)
start_day = pd.to_datetime(start_day_str).date()

start_period = pd.to_datetime(str(start_day)) + pd.Timedelta(hours=7)
end_period = start_period + pd.Timedelta(hours=24)
print(f"📅 Analysis period: {start_period} → {end_period}")

# --------------------------
# 💡 Choose light cycle type
light_cycle = simpledialog.askstring(
    "Light Cycle Selection",
    "Choose the type of light cycle:\n"
    "1 = LD1:1 --> Alternating 1h light / 1h dark\n"
    "2 = DD --> 24h dark\n"
    "3 = LD 12:12 --> 12h light (7–19h) / 12h dark (19–7h)\n"
    "(Enter 1, 2 or 3)"
)

if light_cycle not in ["1", "2", "3"]:
    raise ValueError("❌ Invalid choice. Restart the script and enter 1, 2, or 3.")

# --------------------------
# 📁 Output folder
output_root = r"C:\Users\pablo\OneDrive\Bureau\Program Output"
base_name = os.path.splitext(os.path.basename(file_path))[0]
output_dir = os.path.join(output_root, f"{base_name}_{start_day}_LD11_7h_7h")
os.makedirs(output_dir, exist_ok=True)
print(f"📁 Output folder: {output_dir}")

# --------------------------
# 📊 Read Excel file
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

# --------------------------
# 🧹 Data cleaning
df = df[pd.to_numeric(df["Animal"], errors="coerce").notna()]
df["Animal"] = df["Animal"].astype(int)
df["DateTime"] = pd.to_datetime(df["Date"].astype(str).str.strip() + " " + df["Time"].astype(str).str.strip(), errors="coerce")

for col in ["RER", "XT_YT", "Feed"]:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

# Energy Expenditure (EE)
if "Unnamed: 16" in df.columns:
    df["EE"] = pd.to_numeric(df["Unnamed: 16"], errors="coerce")

df = df.sort_values(["Animal", "DateTime"]).copy()
df["Feed_diff"] = df.groupby("Animal")["Feed"].diff()
df.loc[df["Feed_diff"] < 0, "Feed_diff"] = 0
df["XT_YT"] = df["XT_YT"] / 8000

# --------------------------
# 🧮 Select period 7 AM → 7 AM next day
df_day = df[(df["DateTime"] >= start_period) & (df["DateTime"] < end_period)].copy()

# --------------------------
# 💾 Export full data (no averaging)
output_file = os.path.join(output_dir, f"{base_name}_{start_day}_LD11_7h_7h_raw.xlsx")
df_day.to_excel(output_file, index=False)
print(f"✅ Raw data exported: {output_file}")

# --------------------------
# ☀️🌙 Light cycle visualization
def add_light_cycle(ax, day, cycle_type):
    start = pd.to_datetime(str(day)) + pd.Timedelta(hours=7)

    if cycle_type == "1":
        # Alternating 1h light / 1h dark
        for h in range(0, 24, 2):
            night_start = start + pd.Timedelta(hours=h + 1)
            night_end = start + pd.Timedelta(hours=h + 2)
            ax.axvspan(night_start, night_end, color='gray', alpha=0.2)

    elif cycle_type == "2":
        # 24h dark
        ax.axvspan(start, start + pd.Timedelta(hours=24), color='gray', alpha=0.3)

    elif cycle_type == "3":
        # 12h light (7–19h) / 12h dark (19–7h next day)
        night_start = start + pd.Timedelta(hours=12)   # 19h same day
        night_end = night_start + pd.Timedelta(hours=12)  # 7h next day
        ax.axvspan(night_start, night_end, color='gray', alpha=0.3)

# --------------------------
# 📈 Multi-axis individual graphs (using all data points)
animals = df_day["Animal"].unique()
for animal in animals:
    fig, ax1 = plt.subplots(figsize=(14, 6))
    add_light_cycle(ax1, start_day, light_cycle)

    df_animal = df_day[df_day["Animal"] == animal]

    # Axis 1: RER
    if "RER" in df_animal.columns:
        ax1.plot(df_animal["DateTime"], df_animal["RER"],
                 color='blue', marker='o', linestyle='-', linewidth=1, markersize=3, label="RER")
    ax1.set_xlabel("Hour")
    ax1.set_ylabel("RER", color='blue')
    ax1.tick_params(axis='y', labelcolor='blue')
    ax1.xaxis.set_major_locator(mdates.HourLocator(interval=2))
    ax1.xaxis.set_major_formatter(mdates.DateFormatter('%Hh'))

    # Axis 2: XT+YT
    ax2 = ax1.twinx()
    if "XT_YT" in df_animal.columns:
        ax2.plot(df_animal["DateTime"], df_animal["XT_YT"],
                 color='red', marker='s', linestyle='-', linewidth=1, markersize=3, alpha=0.7, label="XT+YT / 8000")
    ax2.set_ylabel("XT+YT / 8000", color='red')
    ax2.tick_params(axis='y', labelcolor='red')

    # Axis 3: Feed
    ax3 = ax1.twinx()
    if "Feed_diff" in df_animal.columns:
        ax3.plot(df_animal["DateTime"], df_animal["Feed_diff"],
                 color='green', marker='D', linestyle='-', linewidth=1.5, markersize=3, label="Feed (g)")
    ax3.set_ylabel("Feed (g)", color='green')
    ax3.tick_params(axis='y', labelcolor='green')
    ax3.spines['right'].set_position(('outward', 60))

    # Axis 4: EE
    ax4 = ax1.twinx()
    if "EE" in df_animal.columns:
        ax4.plot(df_animal["DateTime"], df_animal["EE"],
                 color='#800080', marker='^', linestyle='-', linewidth=1.5, markersize=3, label="EE (kcal)")
    ax4.set_ylabel("EE (kcal)", color='#800080')
    ax4.tick_params(axis='y', labelcolor='#800080')
    ax4.spines['right'].set_position(('outward', 120))

    ax1.set_title(f"Animal {animal} - {start_day} (Cycle {light_cycle})")
    fig.legend(loc="upper left", bbox_to_anchor=(0.1, 0.9))
    ax1.grid(True, axis='y', linestyle='--', alpha=0.7)
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, f"Graph_Animal{animal}_{start_day}_Cycle{light_cycle}_raw.png"))
    plt.close()

print("✅ Multi-axis graphs successfully generated (raw data)")

# --------------------------
# 📈 Individual metric graphs (no averaging)
for animal in animals:
    df_animal = df_day[df_day["Animal"] == animal]
    for metric, color, ylabel, marker in [
        ("RER", "blue", "RER", "o"),
        ("XT_YT", "red", "XT+YT / 8000", "s"),
        ("Feed_diff", "green", "Feed (g)", "D"),
        ("EE", "#800080", "EE (kcal)", "^")
    ]:
        if metric in df_animal.columns:
            fig, ax = plt.subplots(figsize=(14, 6))
            add_light_cycle(ax, start_day, light_cycle)
            ax.plot(df_animal["DateTime"], df_animal[metric],
                    color=color, marker=marker, linestyle='-', linewidth=1, markersize=3)
            ax.set_title(f"Animal {animal} - {metric} - {start_day} (Cycle {light_cycle})")
            ax.set_xlabel("Hour")
            ax.set_ylabel(ylabel, color=color)
            ax.tick_params(axis='y', labelcolor=color)
            ax.xaxis.set_major_locator(mdates.HourLocator(interval=2))
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%Hh'))
            ax.grid(True, linestyle='--', alpha=0.7)
            plt.xticks(rotation=45)
            plt.tight_layout()
            plt.savefig(os.path.join(output_dir, f"Graph_Animal{animal}_{metric}_{start_day}_Cycle{light_cycle}_raw.png"))
            plt.close()

print("✅ Individual metric graphs successfully generated (raw data)")

print(f"\n📦 All output files are located in: {output_dir}")
