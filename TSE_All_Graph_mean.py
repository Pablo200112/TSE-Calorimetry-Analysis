# -*- coding: utf-8 -*-
"""
Created on Thu Oct  9 09:34:28 2025
@author: pablo
"""

import os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from tkinter import Tk, filedialog, simpledialog, messagebox

# --------------------------
# 📂 Selecting the Excel file (.xlsx)
Tk().withdraw()
file_path = filedialog.askopenfilename(
    title="Select your merged Excel file",
    filetypes=[("Excel Files", "*.xlsx *.xls")]
)
if not file_path:
    raise FileNotFoundError("❌ No file selected. Restart the script and select an Excel file.")
print(f"✅ Selected file : {file_path}")

# --------------------------
# 📁 Output directory
output_root = r"D:\pablo.SAIDI\Desktop\Sortie programme calo"
base_name = os.path.splitext(os.path.basename(file_path))[0]
output_dir = os.path.join(output_root, base_name)
os.makedirs(output_dir, exist_ok=True)
print(f"📁 Output folder : {output_dir}")

# --------------------------
# 📊 Reading the Excel file
df = pd.read_excel(file_path, sheet_name='PS 2025 02')
df.columns = df.columns.str.strip()
print("🧾 Detected columns:", df.columns.tolist())

# 🧱 Main renaming
df = df.rename(columns={
    "PS 2025 02": "Date",
    "Unnamed: 1": "Time",
    "TX002": "Animal",
    "Unnamed: 13": "RER",
    "Unnamed: 14": "XT_YT",
    "Unnamed: 15": "Feed"
})

# 🔍 Forcing the Energy Expenditure column (column Q)
if len(df.columns) >= 17:
    ee_col_name = df.columns[16]
    df = df.rename(columns={ee_col_name: "EE"})
    print(f"✅ Forced Energy Expenditure column : {ee_col_name} (column Q)")
else:
    print("⚠️ The file doesn't contain a Q column. Please check the file format.")
    df["EE"] = None

# --------------------------
# Relevant columns
useful_columns = ["Date", "Time", "Animal", "RER", "XT_YT", "Feed", "EE"]
df = df[[c for c in useful_columns if c in df.columns]].copy()

# --------------------------
# Cleaning and formatting
df = df[pd.to_numeric(df["Animal"], errors="coerce").notna()]
df["Animal"] = df["Animal"].astype(int)
df["DateTime"] = pd.to_datetime(df["Date"].astype(str) + " " + df["Time"].astype(str), errors="coerce")

for col in ["RER", "XT_YT", "Feed", "EE"]:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

df = df.sort_values(["Animal", "DateTime"])

# --------------------------
# Differential Feed
df["Feed_diff"] = df.groupby("Animal")["Feed"].diff()
df.loc[df["Feed_diff"] < 0, "Feed_diff"] = 0

# ❓ Ask user if values >2 should be excluded
root = Tk()
root.withdraw()
exclude_feed_outliers = messagebox.askyesno(
    "Feed_diff Filtering",
    "Do you want to exclude Feed_diff values greater than 2 ?"
)
root.destroy()

if exclude_feed_outliers:
    print("⛔ Excluding Feed_diff values > 2")
    df.loc[df["Feed_diff"] > 2, "Feed_diff"] = None
else:
    print("✔ Keeping all Feed_diff values (no filtering)")

# --------------------------
# Normalizing XT_YT
df["XT_YT"] = df["XT_YT"] / 8000

# Day / Hour
df["Day"] = df["DateTime"].dt.date
df["Hour"] = df["DateTime"].dt.hour

# --------------------------
# Hourly averages per animal
rer_pivot = df.pivot_table(index=["Day", "Hour"], columns="Animal", values="RER", aggfunc="mean")
xtyt_pivot = df.pivot_table(index=["Day", "Hour"], columns="Animal", values="XT_YT", aggfunc="sum")
feed_pivot = df.pivot_table(index=["Day", "Hour"], columns="Animal", values="Feed_diff", aggfunc="sum")
ee_pivot = df.pivot_table(index=["Day", "Hour"], columns="Animal", values="EE", aggfunc="mean")

rer_pivot.columns = [f"RER_Animal{c}" for c in rer_pivot.columns]
xtyt_pivot.columns = [f"XT_YT_Animal{c}" for c in xtyt_pivot.columns]
feed_pivot.columns = [f"Feed_Animal{c}" for c in feed_pivot.columns]
ee_pivot.columns = [f"EE_Animal{c}" for c in ee_pivot.columns]

df_pivot = pd.concat([rer_pivot, xtyt_pivot, feed_pivot, ee_pivot], axis=1).reset_index()
df_pivot["DateTime"] = pd.to_datetime(df_pivot["Day"].astype(str)) + pd.to_timedelta(df_pivot["Hour"], unit='h')

# --------------------------
# Export to Excel
output_file = os.path.join(output_dir, f"{base_name}_Hourly_Averages_per_Animal.xlsx")
df_pivot.to_excel(output_file, index=False)
print("✅ File exported:", output_file)

# --------------------------
# 🌙 Day/Night Cycle
def add_night_zones(ax, days):
    for day in days:
        night_start = pd.to_datetime(str(day) + " 19:00")
        night_end = pd.to_datetime(str(day) + " 23:59:59")
        ax.axvspan(night_start, night_end, color='gray', alpha=0.2)
        next_day = pd.to_datetime(str(day) + " 00:00") + pd.Timedelta(days=1)
        morning_end = next_day + pd.Timedelta(hours=7)
        ax.axvspan(next_day, morning_end, color='gray', alpha=0.2)

# --------------------------
# 🌗 Special Cycles
def add_alternation_cycle(ax, day, start_hour=7):
    start = pd.to_datetime(str(day) + f" {start_hour}:00")
    end = start + pd.Timedelta(hours=24)
    hours = pd.date_range(start=start, end=end, freq="1H")
    for i in range(len(hours)-1):
        if i % 2 == 1:
            ax.axvspan(hours[i], hours[i+1], color='gray', alpha=0.3)

def add_darkness_cycle(ax, day, start_hour=7):
    start = pd.to_datetime(str(day) + f" {start_hour}:00")
    ax.axvspan(start, start + pd.Timedelta(hours=24), color='black', alpha=0.25)

# --------------------------
# 🪟 Window to select special days
root = Tk()
root.withdraw()
alternation_day = simpledialog.askstring("Alternation Day", "📅 Date with LD1:1 alternation (YYYY-MM-DD) :")
darkness_day = simpledialog.askstring("Darkness Day", "🌑 Date with full darkness (YYYY-MM-DD) :")
root.destroy()
print(f"🌗 Alternation: {alternation_day} | 🌑 Darkness: {darkness_day}")

# --------------------------
# Individual Graphs
animals = df["Animal"].unique()
for animal in animals:
    fig, ax1 = plt.subplots(figsize=(14, 6))

    for day in df_pivot["Day"].unique():
        if alternation_day and str(day) == alternation_day:
            add_alternation_cycle(ax1, alternation_day)
        elif darkness_day and str(day) == darkness_day:
            add_darkness_cycle(ax1, darkness_day)
        else:
            add_night_zones(ax1, [day])

    if f"RER_Animal{animal}" in df_pivot.columns:
        ax1.scatter(df_pivot["DateTime"], df_pivot[f"RER_Animal{animal}"], label="RER", color='blue', s=15)
    if f"XT_YT_Animal{animal}" in df_pivot.columns:
        ax1.bar(df_pivot["DateTime"], df_pivot[f"XT_YT_Animal{animal}"], width=0.03, color='red', alpha=0.6, label="XT_YT [u.a]")
    if f"EE_Animal{animal}" in df_pivot.columns:
        ax1.plot(df_pivot["DateTime"], df_pivot[f"EE_Animal{animal}"], color='purple', linewidth=2, label="EE [kcal/h]")

    ax1.set_xlabel("Date and Hour", fontsize=14, fontweight='bold')
    ax1.set_ylabel("RER / XT+YT / EE", fontsize=14, fontweight='bold')
    ax1.xaxis.set_major_locator(mdates.HourLocator(byhour=[0, 12]))
    ax1.xaxis.set_major_formatter(mdates.DateFormatter('%d-%Hh'))
    fig.autofmt_xdate(rotation=45, ha='right')

    ax2 = ax1.twinx()
    if f"Feed_Animal{animal}" in df_pivot.columns:
        ax2.plot(df_pivot["DateTime"], df_pivot[f"Feed_Animal{animal}"], color='green', linewidth=2, label="Feed [g/h]")
    ax2.set_ylabel("Hourly Feed", color='green', fontsize=14, fontweight='bold')

    ax1.set_title(f"Animal {animal} : RER, XT+YT, EE, Feed", fontsize=16, fontweight='bold')
    fig.legend(loc="upper left", bbox_to_anchor=(0.1, 0.9))
    ax1.grid(True, axis='y')
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, f"Graph_Animal{animal}_RER_XT_YT_EE_Feed.png"))
    plt.close()

print("✅ Individual graphs generated successfully")

# --------------------------
# Global Graphs
def generate_global_graph(df_pivot, animals, metric_prefix, title, ylabel, filename):
    fig, ax = plt.subplots(figsize=(14, 6))

    for day in df_pivot["Day"].unique():
        if alternation_day and str(day) == alternation_day:
            add_alternation_cycle(ax, alternation_day)
        elif darkness_day and str(day) == darkness_day:
            add_darkness_cycle(ax, darkness_day)
        else:
            add_night_zones(ax, [day])

    for animal in animals:
        col = f"{metric_prefix}_Animal{animal}"
        if col in df_pivot.columns:
            ax.plot(df_pivot["DateTime"], df_pivot[col], label=f"Animal {animal}")

    ax.set_title(title, fontsize=16, fontweight='bold')
    ax.set_xlabel("Date and Hour", fontsize=14, fontweight='bold')
    ax.set_ylabel(ylabel, fontsize=14, fontweight='bold')
    ax.xaxis.set_major_locator(mdates.HourLocator(byhour=[0, 12]))
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%d-%Hh'))
    fig.autofmt_xdate(rotation=45, ha='right')
    ax.legend()
    ax.grid(True)
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, filename))
    plt.close()

# Global graphs
generate_global_graph(df_pivot, animals, "RER", "Average RER - All animals", "RER (hourly average)", "Graph_Global_RER.png")
generate_global_graph(df_pivot, animals, "XT_YT", "Average XT+YT - All animals", "XT+YT (hourly average)", "Graph_Global_XT_YT.png")
generate_global_graph(df_pivot, animals, "Feed", "Hourly Feed - All animals", "Hourly Feed (g/h)", "Graph_Global_Feed.png")

ee_cols = [col for col in df_pivot.columns if col.startswith("EE_Animal")]
if ee_cols:
    generate_global_graph(df_pivot, animals, "EE", "Average Energy Expenditure - All animals", "EE [kcal/h]", "Graph_Global_EE.png")

print("✅ Global graphs for RER, XT+YT, Feed, and EE generated successfully")
print(f"\n📦 All files are in: {output_dir}")

