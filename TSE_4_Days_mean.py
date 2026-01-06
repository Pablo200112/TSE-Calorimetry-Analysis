# -*- coding: utf-8 -*-
"""
Complete Script: Calorimetry Analysis LD11
Created by Pablo SAIDI
(Version: raw timestamps shifted BEFORE averaging: choice = BEGIN / CENTER / END)
"""

import os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from tkinter import Tk, filedialog, simpledialog, messagebox

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
# 🕒 Choose how to align raw sampling windows
shift_choice = simpledialog.askstring(
   "Timestamp Position",
    "Sampling window is 15 min.\n"
    "Choose how to position each data point:\n\n"
    "1 = beginning of window (ex: 08:00 → 07:45)\n"
    "2 = center of window, recommended (ex: 08:00 → 07:52:30)\n"
    "3 = end of window (no correction)\n\n"
    "Enter 1, 2, or 3:"
)

if shift_choice not in ["1", "2", "3"]:
    raise ValueError("❌ Invalid choice. Restart the script and enter 1, 2 or 3.")

# Mapping: we subtract offset (because input timestamps are at window end)
if shift_choice == "1":
    offset_minutes = 15.0
    print("⏱️ Alignment: BEGINNING of window (timestamps shifted -15 minutes).")
elif shift_choice == "2":
    offset_minutes = 7.5
    print("⏱️ Alignment: CENTER of window (timestamps shifted -7.5 minutes).")
else:
    offset_minutes = 0.0
    print("⏱️ Alignment: END of window (no timestamp shift).")

# --------------------------
# 💡 Choose light cycle type
light_cycle = simpledialog.askstring(
    "Light Cycle Selection",
    "Choose the type of light cycle:\n"
    "1 = LD1:1 --> Alternating 1h light / 1h dark\n"
    "2 = DD --> 24h dark\n"
    "3 = LD 12:12 --> 12h light / 12h dark\n"
    "(Enter 1, 2 or 3)"
)

if light_cycle not in ["1", "2", "3"]:
    raise ValueError("❌ Invalid choice. Restart the script and enter 1, 2, or 3.")

# --------------------------
# 📁 Output folder
output_root = r"D:\pablo.SAIDI\Desktop\Sortie programme calo"
base_name = os.path.splitext(os.path.basename(file_path))[0]
output_dir = os.path.join(output_root, f"{base_name}_{start_day}_LD11_7h_7h")
os.makedirs(output_dir, exist_ok=True)
print(f"📁 Output folder: {output_dir}")

# --------------------------
# 📊 Read Excel file
df = pd.read_excel(file_path, sheet_name='2em PS 2025 02')
df.columns = df.columns.str.strip()

df = df.rename(columns={
    "2em PS 2025 02": "Date",
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
df["DateTime"] = pd.to_datetime(
    df["Date"].astype(str).str.strip() + " " + df["Time"].astype(str).str.strip(),
    errors="coerce"
)

for col in ["RER", "XT_YT", "Feed"]:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

# Energy Expenditure (EE)
if "Unnamed: 16" in df.columns:
    df["EE"] = pd.to_numeric(df["Unnamed: 16"], errors="coerce")

df = df.sort_values(["Animal", "DateTime"]).copy()

# Compute Feed differences
df["Feed_diff"] = df.groupby("Animal")["Feed"].diff()
df.loc[df["Feed_diff"] < 0, "Feed_diff"] = 0

# --------------------------
# 🧪 Option to exclude Feed_diff > 2 g
root = Tk()
root.withdraw()
exclude_feed = messagebox.askyesno(
    "Feed_diff > 2 g",
    "Do you want to exclude Feed_diff values greater than 2 g?"
)
root.destroy()

if exclude_feed:
    print("⛔ Excluding Feed_diff values > 2 g")
    df["Feed_diff"] = df["Feed_diff"].where(df["Feed_diff"] <= 2, None)
else:
    print("✔ Keeping all Feed_diff values (no filtering)")

# --------------------------
# Normalize XT+YT
df["XT_YT"] = df["XT_YT"] / 8000

# --------------------------
# ⏱️ SHIFT RAW TIMESTAMPS
df["DateTime_shifted"] = df["DateTime"] - pd.to_timedelta(offset_minutes, unit="m")
df = df.sort_values(["Animal", "DateTime_shifted"]).copy()

# --------------------------
# 🧮 Select period 7 AM → 7 AM next day using shifted timestamps
df_day = df[(df["DateTime_shifted"] >= start_period) & (df["DateTime_shifted"] < end_period)].copy()
df_day["Relative_Hour"] = ((df_day["DateTime_shifted"] - start_period).dt.total_seconds() // 3600).astype(int)

# --------------------------
# 📘 Export shifted raw data
output_file_shifted = os.path.join(output_dir, f"{base_name}_{start_day}_shifted_raw.xlsx")
df_day.to_excel(output_file_shifted, index=False)
print(f"✅ Shifted raw data exported: {output_file_shifted}")

# --------------------------
# 📊 Hourly averages / sums
rer_pivot = df_day.pivot_table(index="Relative_Hour", columns="Animal", values="RER", aggfunc="mean")
xtyt_pivot = df_day.pivot_table(index="Relative_Hour", columns="Animal", values="XT_YT", aggfunc="sum")
feed_pivot = df_day.pivot_table(index="Relative_Hour", columns="Animal", values="Feed_diff", aggfunc="sum")

rer_pivot.columns = [f"RER_Animal{col}" for col in rer_pivot.columns]
xtyt_pivot.columns = [f"XT_YT_Animal{col}" for col in xtyt_pivot.columns]
feed_pivot.columns = [f"Feed_Animal{col}" for col in feed_pivot.columns]

if "EE" in df_day.columns:
    ee_pivot = df_day.pivot_table(index="Relative_Hour", columns="Animal", values="EE", aggfunc="sum")
    ee_pivot.columns = [f"EE_Animal{col}" for col in ee_pivot.columns]
    df_pivot = pd.concat([rer_pivot, xtyt_pivot, feed_pivot, ee_pivot], axis=1).reset_index()
else:
    df_pivot = pd.concat([rer_pivot, xtyt_pivot, feed_pivot], axis=1).reset_index()

df_pivot["DateTime"] = start_period + pd.to_timedelta(df_pivot["Relative_Hour"], unit='h') + pd.to_timedelta(0.5, unit='h')

# --------------------------
# 💾 Export hourly pivot
output_file = os.path.join(output_dir, f"{base_name}_{start_day}_LD11_7h_7h.xlsx")
df_pivot.to_excel(output_file, index=False)
print(f"✅ Hourly pivot exported: {output_file}")

# --------------------------
# ☀️🌙 Light cycle visualization
def add_light_cycle(ax, day, cycle_type):
    start = pd.to_datetime(str(day)) + pd.Timedelta(hours=7)
    if cycle_type == "1":
        for h in range(0, 24, 2):
            ax.axvspan(start + pd.Timedelta(hours=h+1), start + pd.Timedelta(hours=h+2), color='gray', alpha=0.2)
    elif cycle_type == "2":
        ax.axvspan(start, start + pd.Timedelta(hours=24), color='gray', alpha=0.3)
    elif cycle_type == "3":
        night_start = start + pd.Timedelta(hours=12)
        night_end = night_start + pd.Timedelta(hours=12)
        ax.axvspan(night_start, night_end, color='gray', alpha=0.3)

# --------------------------
# 📈 Multi-axis individual graphs
animals = df_day["Animal"].unique()
for animal in animals:
    fig, ax1 = plt.subplots(figsize=(14, 6))
    add_light_cycle(ax1, start_day, light_cycle)

    if f"RER_Animal{animal}" in df_pivot.columns:
        ax1.plot(df_pivot["DateTime"], df_pivot[f"RER_Animal{animal}"],
                 color='blue', marker='o', linestyle='-', linewidth=1.5, markersize=5, label="RER")
    ax1.set_xlabel("Hour")
    ax1.set_ylabel("RER", color='blue')
    ax1.tick_params(axis='y', labelcolor='blue')
    ax1.xaxis.set_major_locator(mdates.HourLocator(interval=2))
    ax1.xaxis.set_major_formatter(mdates.DateFormatter('%Hh'))

    ax2 = ax1.twinx()
    if f"XT_YT_Animal{animal}" in df_pivot.columns:
        ax2.plot(df_pivot["DateTime"], df_pivot[f"XT_YT_Animal{animal}"],
                 color='red', marker='s', linestyle='-', linewidth=1.5, markersize=5, alpha=0.7, label="XT+YT / 8000")
    ax2.set_ylabel("XT+YT / 8000", color='red')
    ax2.tick_params(axis='y', labelcolor='red')

    ax3 = ax1.twinx()
    if f"Feed_Animal{animal}" in df_pivot.columns:
        ax3.plot(df_pivot["DateTime"], df_pivot[f"Feed_Animal{animal}"],
                 color='green', marker='D', linestyle='-', linewidth=2, markersize=4, label="Feed (g/h)")
    ax3.set_ylabel("Feed (g/h)", color='green')
    ax3.tick_params(axis='y', labelcolor='green')
    ax3.spines['right'].set_position(('outward', 60))

    ax4 = ax1.twinx()
    if f"EE_Animal{animal}" in df_pivot.columns:
        ax4.plot(df_pivot["DateTime"], df_pivot[f"EE_Animal{animal}"],
                 color='#800080', marker='^', linestyle='-', linewidth=2, markersize=4, label="EE (kcal/h)")
    ax4.set_ylabel("EE (kcal/h)", color='#800080')
    ax4.tick_params(axis='y', labelcolor='#800080')
    ax4.spines['right'].set_position(('outward', 120))

    ax1.set_title(f"Animal {animal} - {start_day} (Cycle {light_cycle})")
    fig.legend(loc="upper left", bbox_to_anchor=(0.1, 0.9))
    ax1.grid(True, axis='y', linestyle='--', alpha=0.7)
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, f"Graph_Animal{animal}_{start_day}_Cycle{light_cycle}.png"))
    plt.close()

print("✅ Multi-axis graphs successfully generated")

# --------------------------
# 📈 Individual metric graphs
for animal in animals:
    for metric, color, ylabel, marker in [
        ("RER", "blue", "RER", "o"),
        ("XT_YT", "red", "XT+YT / 8000", "s"),
        ("Feed", "green", "Feed (g/h)", "D"),
        ("EE", "#800080", "EE (kcal/h)", "^")
    ]:
        col_name = f"{metric}_Animal{animal}"
        if col_name in df_pivot.columns:
            fig, ax = plt.subplots(figsize=(14, 6))
            add_light_cycle(ax, start_day, light_cycle)
            ax.plot(df_pivot["DateTime"], df_pivot[col_name],
                    color=color, marker=marker, linestyle='-', linewidth=1.5, markersize=5)
            ax.set_title(f"Animal {animal} - {metric} - {start_day} (Cycle {light_cycle})")
            ax.set_xlabel("Hour")
            ax.set_ylabel(ylabel, color=color)
            ax.tick_params(axis='y', labelcolor=color)
            ax.xaxis.set_major_locator(mdates.HourLocator(interval=2))
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%Hh'))
            ax.grid(True, linestyle='--', alpha=0.7)
            plt.xticks(rotation=45)
            plt.tight_layout()
            plt.savefig(os.path.join(output_dir, f"Graph_Animal{animal}_{metric}_{start_day}_Cycle{light_cycle}.png"))
            plt.close()

print("✅ Individual metric graphs successfully generated")

# --------------------------
# 📊 Global graphs
def generate_global_graph(df_pivot, animals, metric_prefix, title, ylabel, filename, color='blue', marker='o'):
    fig, ax = plt.subplots(figsize=(14, 6))
    add_light_cycle(ax, start_day, light_cycle)

    for animal in animals:
        col = f"{metric_prefix}_Animal{animal}"
        if col in df_pivot.columns:
            ax.plot(df_pivot["DateTime"], df_pivot[col],
                    color=color, marker=marker, linestyle='-', linewidth=1.5, markersize=5, label=f"Animal {animal}")

    ax.set_title(f"{title} - {start_day} (Cycle {light_cycle})")
    ax.set_xlabel("Hour")
    ax.set_ylabel(ylabel)
    ax.xaxis.set_major_locator(mdates.HourLocator(interval=2))
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Hh'))
    ax.legend()
    ax.grid(True, linestyle='--', alpha=0.7)
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, filename))
    plt.close()

# 🔹 Generate global graphs
generate_global_graph(df_pivot, animals, "RER", "Average RER per hour - All animals",
                      "RER (hourly average)", f"Graph_Global_RER_All_Animals_{start_day}_Cycle{light_cycle}.png", color='blue')
generate_global_graph(df_pivot, animals, "XT_YT", "Average XT+YT/8000 per hour - All animals",
                      "XT+YT / 8000", f"Graph_Global_XT_YT_All_Animals_{start_day}_Cycle{light_cycle}.png", color='red')
generate_global_graph(df_pivot, animals, "Feed", "Hourly Feed - All animals",
                      "Hourly Feed", f"Graph_Global_Feed_All_Animals_{start_day}_Cycle{light_cycle}.png", color='green')
generate_global_graph(df_pivot, animals, "EE", "Hourly EE - All animals",
                      "Hourly EE", f"Graph_Global_EE_All_Animals_{start_day}_Cycle{light_cycle}.png", color='#800080', marker='^')

print("✅ All graphs successfully generated")
print(f"\n📦 All output files are located in: {output_dir}")
