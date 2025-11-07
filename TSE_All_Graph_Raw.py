# -*- coding: utf-8 -*-
"""
Created on Thu Oct  9 09:34:28 2025
Modified for raw 15-min data + optional 1h rolling mean smoothing
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
output_root = r"C:\Users\pablo\OneDrive\Bureau\Program Output"
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

# Differential Feed
df["Feed_diff"] = df.groupby("Animal")["Feed"].diff()
df.loc[df["Feed_diff"] < 0, "Feed_diff"] = 0

# Normalizing XT_YT
df["XT_YT"] = df["XT_YT"] / 8000

# Day / Hour
df["Day"] = df["DateTime"].dt.date
df["Hour"] = df["DateTime"].dt.hour

# --------------------------
# ❓ Ask user if smoothing is desired
root = Tk()
root.withdraw()
apply_smoothing = messagebox.askyesno(
    "Rolling Mean",
    "Do you want to smooth the data with a 1-hour rolling mean (4 points of 15 min)?"
)
root.destroy()

if apply_smoothing:
    print("🔄 Applying 1-hour rolling mean smoothing (4x15min)...")
    df[["RER", "XT_YT", "EE", "Feed_diff"]] = (
        df.groupby("Animal")[["RER", "XT_YT", "EE", "Feed_diff"]].transform(lambda x: x.rolling(window=4, min_periods=1).mean())
    )
else:
    print("🚫 No smoothing applied (raw 15-min data used).")

# --------------------------
# Export raw (or smoothed) 15-min data
suffix = "_Smoothed" if apply_smoothing else "_Raw"
output_file = os.path.join(output_dir, f"{base_name}{suffix}_15min_per_Animal.xlsx")
df.to_excel(output_file, index=False)
print("✅ 15-min data exported:", output_file)

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
        t1, t2 = hours[i], hours[i+1]
        if i % 2 == 1:
            ax.axvspan(t1, t2, color='gray', alpha=0.3)

def add_darkness_cycle(ax, day, start_hour=7):
    start = pd.to_datetime(str(day) + f" {start_hour}:00")
    end = start + pd.Timedelta(hours=24)
    ax.axvspan(start, end, color='black', alpha=0.25)

# --------------------------
# 🪟 Window to select special days
root = Tk()
root.withdraw()
alternation_day = simpledialog.askstring("Alternation Day", "📅 Date of the day with 1h/1h alternation (LD1:1) (YYYY-MM-DD):")
darkness_day = simpledialog.askstring("Darkness Day", "🌑 Date of the day with total darkness (DD) (YYYY-MM-DD):")
root.destroy()
print(f"🌗 Alternation: {alternation_day} | 🌑 Darkness: {darkness_day}")

# --------------------------
# Individual Graphs (15-min data)
animals = df["Animal"].unique()
for animal in animals:
    sub = df[df["Animal"] == animal]
    fig, ax1 = plt.subplots(figsize=(14, 6))

    # Conditional display by day
    for day in sub["Day"].unique():
        if alternation_day and str(day) == alternation_day:
            add_alternation_cycle(ax1, alternation_day)
        elif darkness_day and str(day) == darkness_day:
            add_darkness_cycle(ax1, darkness_day)
        else:
            add_night_zones(ax1, [day])

    # Plot data
    if "RER" in sub.columns:
        ax1.plot(sub["DateTime"], sub["RER"], label="RER", color='blue', linewidth=1.5)
    if "XT_YT" in sub.columns:
        ax1.bar(sub["DateTime"], sub["XT_YT"], width=0.01, color='red', alpha=0.6, label="XT_YT [a.u.]")
    if "EE" in sub.columns:
        ax1.plot(sub["DateTime"], sub["EE"], color='purple', linewidth=2, label="EE [kcal/h]")

    ax1.set_xlabel("Date and Hour", fontsize=14, fontweight='bold')
    ax1.set_ylabel("RER / XT+YT / EE", fontsize=14, fontweight='bold')
    ax1.xaxis.set_major_locator(mdates.HourLocator(byhour=[0, 12]))
    ax1.xaxis.set_major_formatter(mdates.DateFormatter('%d-%Hh'))
    fig.autofmt_xdate(rotation=45, ha='right')

    # Feed on secondary axis
    ax2 = ax1.twinx()
    ax2.plot(sub["DateTime"], sub["Feed_diff"], color='green', linewidth=2, label="Feed [g]")
    ax2.set_ylabel("Feed (g per 15 min)", color='green', fontsize=14, fontweight='bold')

    ax1.set_title(f"Animal {animal} : RER, XT+YT, EE, Feed (15-min{' smoothed' if apply_smoothing else ' raw'})", fontsize=16, fontweight='bold')
    fig.legend(loc="upper left", bbox_to_anchor=(0.1, 0.9))
    ax1.grid(True, axis='y')
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, f"Graph_Animal{animal}_15min{suffix}.png"))
    plt.close()

print("✅ Individual 15-min graphs generated successfully")

# --------------------------
# Global Graphs (15-min data)
def generate_global_graph(df, animals, metric_prefix, title, ylabel, filename):
    fig, ax = plt.subplots(figsize=(14, 6))

    for day in df["Day"].unique():
        if alternation_day and str(day) == alternation_day:
            add_alternation_cycle(ax, alternation_day)
        elif darkness_day and str(day) == darkness_day:
            add_darkness_cycle(ax, darkness_day)
        else:
            add_night_zones(ax, [day])

    for animal in animals:
        sub = df[df["Animal"] == animal]
        if metric_prefix in sub.columns:
            ax.plot(sub["DateTime"], sub[metric_prefix], label=f"Animal {animal}")

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
generate_global_graph(df, animals, "RER", f"RER (15-min{' smoothed' if apply_smoothing else ' raw'}) - All animals", "RER", f"Graph_Global_RER{suffix}.png")
generate_global_graph(df, animals, "XT_YT", f"XT+YT (15-min{' smoothed' if apply_smoothing else ' raw'}) - All animals", "XT+YT [a.u.]", f"Graph_Global_XT_YT{suffix}.png")
generate_global_graph(df, animals, "Feed_diff", f"Feed (15-min{' smoothed' if apply_smoothing else ' raw'}) - All animals", "Feed (g/15 min)", f"Graph_Global_Feed{suffix}.png")

if "EE" in df.columns:
    generate_global_graph(df, animals, "EE", f"Energy Expenditure (15-min{' smoothed' if apply_smoothing else ' raw'}) - All animals", "EE [kcal/h]", f"Graph_Global_EE{suffix}.png")

print("✅ Global 15-min graphs generated successfully")
print(f"\n📦 All files are in: {output_dir}")
