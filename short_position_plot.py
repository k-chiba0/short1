import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.ticker import FuncFormatter
from matplotlib import gridspec
from pathlib import Path
import os

# フォルダ設定
data_dir = Path("D:/short/data")
csv_dir = Path("D:/short/csv")
chart_dir = Path("D:/short/Chart")
csv_dir.mkdir(parents=True, exist_ok=True)

# 銘柄コード入力
target_code = input("銘柄コードを入力してください（例：5253）: ").strip()

# 総合計をグラフに含めるか確認
include_total = input("空売り総合計線をグラフに表示しますか？ (y/n): ").strip().lower()

# --- Step 1: .xls -> .xlsx 変換 ---
for file in data_dir.glob("*.xls"):
    try:
        df = pd.read_excel(file, engine="xlrd", header=None)
        new_file = file.with_suffix(".xlsx")
        df.to_excel(new_file, index=False, header=False, engine="openpyxl")
        file.unlink()
        print(f"{file.name} を変換して削除しました")
    except Exception as e:
        print(f"{file.name} の変換中にエラー: {e}")

# --- Step 2: 残高ファイル読み込み ---
records = []
start_row = 8

for file in data_dir.glob("*.xlsx"):
    try:
        df = pd.read_excel(file, engine="openpyxl", header=None)
        data = df.iloc[start_row:, :].dropna(subset=[1, 2, 3, 5, 11])  # B, C, D, F, L列に対応

        for _, row in data.iterrows():
            if str(row[2]).strip() == target_code:  # C列（銘柄コード）
                try:
                    calc_date_row = pd.to_datetime(row[1])  # B列（計算年月日）
                    records.append({
                        "date": calc_date_row,
                        "institution": str(row[5]).strip(),  # F列（機関名）
                        "amount": float(row[11]),           # L列（空売り株数）
                        "stock_name": str(row[3]).strip()[:5]  # D列（銘柄名）
                    })
                except Exception as e:
                    print(f"{file.name} 行 {row.name} で日付変換エラー: {e}")

    except Exception as e:
        print(f"{file.name} の読み込み中にエラー: {e}")

# --- Step 3: データ処理 ---
df_all = pd.DataFrame(records)
grouped = df_all.groupby(["date", "institution"])["amount"].sum().reset_index()
pivot_df = grouped.pivot(index="date", columns="institution", values="amount")

# 残高を前日から補完（直前の値で埋める）
pivot_df = pivot_df.sort_index().ffill().fillna(0)

# 合計列を追加
pivot_df_with_total = pivot_df.copy()
pivot_df_with_total["合計"] = pivot_df_with_total.sum(axis=1)

# CSV出力
csv_path = csv_dir / f"{target_code}_short_positions.csv"
pivot_df_with_total.to_csv(csv_path, encoding="utf-8-sig")
print(f"CSV出力完了: {csv_path}")

# --- Step 4: チャートデータの読み込み ---
chart_files = list(chart_dir.glob(f"TimeChart({target_code})*.csv"))
if not chart_files:
    print("ローソク足データが見つかりません。")
    exit()

chart_data = pd.concat([
    pd.read_csv(f, encoding="utf-8", skiprows=1, usecols=[0, 1, 2, 3, 4, 9],
                names=["date", "open", "high", "low", "close", "volume"])
    for f in chart_files
])

# 日付と数値形式の整形
chart_data["date"] = pd.to_datetime(chart_data["date"])
for col in ["open", "high", "low", "close", "volume"]:
    chart_data[col] = chart_data[col].replace(",", "", regex=True).astype(float)

chart_data = chart_data.sort_values("date")

# --- Step 5: グラフ描画 ---
fig = plt.figure(figsize=(14, 10))
gs = gridspec.GridSpec(3, 1, height_ratios=[2, 1, 1])

# 1: 空売り残高推移
ax1 = fig.add_subplot(gs[0])

# 総合計を表示するかどうかで分岐
columns_to_plot = pivot_df_with_total.columns if include_total == 'y' else pivot_df.columns

for col in columns_to_plot:
    ax1.plot(pivot_df_with_total.index, pivot_df_with_total[col], label=col)

ax1.set_title(f"空売り残高推移 - {target_code} ({df_all.iloc[0]['stock_name']})", fontname="MS Gothic")
ax1.set_ylabel("空売り残高（株）", fontname="MS Gothic")
ax1.legend(loc='upper left', bbox_to_anchor=(1, 1), fontsize=8)
ax1.yaxis.set_major_formatter(FuncFormatter(lambda x, _: f"{int(x):,}"))
ax1.grid(True)

# 2: ローソク足チャート
ax2 = fig.add_subplot(gs[1], sharex=ax1)
for _, row in chart_data.iterrows():
    color = 'red' if row['close'] >= row['open'] else 'blue'
    ax2.plot([row['date'], row['date']], [row['low'], row['high']], color=color)
    ax2.plot([row['date'], row['date']], [row['open'], row['close']], color=color, linewidth=6)
ax2.set_ylabel("株価", fontname="MS Gothic")
ax2.grid(True)

# 3: 出来高チャート
ax3 = fig.add_subplot(gs[2], sharex=ax1)
ax3.bar(chart_data["date"], chart_data["volume"], color="gray", width=0.8)
ax3.set_ylabel("出来高", fontname="MS Gothic")
ax3.yaxis.set_major_formatter(FuncFormatter(lambda x, _: f"{int(x):,}"))
ax3.grid(True)

plt.xlabel("日付", fontname="MS Gothic")
plt.tight_layout()
plt.show()