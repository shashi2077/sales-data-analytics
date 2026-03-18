"""
Sales Data Analytics Dashboard
Author: Shashikant Yadav
Description: Performs data cleaning, transformation, and visualization of sales data.
Tech Used: Python, Pandas, Matplotlib, Excel
"""

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import os

# ─────────────────────────────────────────
# 1. Load & Clean Data
# ─────────────────────────────────────────
df = pd.read_csv("sales_data.csv")

print("=" * 55)
print("      SALES DATA ANALYTICS DASHBOARD")
print("=" * 55)
print(f"\nDataset: {df.shape[0]} rows x {df.shape[1]} columns")

# Data Cleaning
print("\n--- Data Cleaning ---")
print(f"Missing values : {df.isnull().sum().sum()}")
print(f"Duplicate rows : {df.duplicated().sum()}")
df.dropna(inplace=True)
df.drop_duplicates(inplace=True)
print("✅ Data is clean and ready for analysis")

# ─────────────────────────────────────────
# 2. Transformation
# ─────────────────────────────────────────
df["Return_Rate_%"]   = (df["Returns"] / df["Sales_Units"] * 100).round(2)
df["ROI_%"]           = ((df["Net_Revenue"] - df["Marketing_Spend"]) / df["Marketing_Spend"] * 100).round(2)

# Monthly aggregation
monthly = df.groupby("Month").agg(
    Total_Revenue    = ("Revenue", "sum"),
    Net_Revenue      = ("Net_Revenue", "sum"),
    Total_Units      = ("Sales_Units", "sum"),
    Total_Returns    = ("Returns", "sum"),
    Marketing_Spend  = ("Marketing_Spend", "sum")
).reset_index()
monthly["Profit_Margin_%"] = (monthly["Net_Revenue"] / monthly["Total_Revenue"] * 100).round(2)
monthly["MoM_Growth_%"]    = monthly["Total_Revenue"].pct_change().mul(100).round(2)

# Product-wise aggregation
product = df.groupby("Product").agg(
    Total_Revenue = ("Revenue", "sum"),
    Total_Units   = ("Sales_Units", "sum"),
    Avg_ROI       = ("ROI_%", "mean")
).reset_index()

# Region-wise aggregation
region = df.groupby("Region").agg(
    Total_Revenue = ("Revenue", "sum"),
    Total_Units   = ("Sales_Units", "sum")
).reset_index()

# ─────────────────────────────────────────
# 3. KPIs
# ─────────────────────────────────────────
total_rev    = df["Revenue"].sum()
total_net    = df["Net_Revenue"].sum()
total_units  = df["Sales_Units"].sum()
total_ret    = df["Returns"].sum()
avg_roi      = df["ROI_%"].mean()
best_product = product.loc[product["Total_Revenue"].idxmax(), "Product"]
best_region  = region.loc[region["Total_Revenue"].idxmax(), "Region"]
best_month   = monthly.loc[monthly["Total_Revenue"].idxmax(), "Month"]

print("\n--- Key Performance Indicators ---")
print(f"Total Revenue        : ₹{total_rev:,.0f}")
print(f"Total Net Revenue    : ₹{total_net:,.0f}")
print(f"Total Units Sold     : {total_units:,}")
print(f"Total Returns        : {total_ret:,}")
print(f"Average ROI          : {avg_roi:.2f}%")
print(f"Best Product         : {best_product}")
print(f"Best Region          : {best_region}")
print(f"Best Month           : {best_month}")

print("\n--- Monthly Trend ---")
print(monthly[["Month","Total_Revenue","Net_Revenue","Total_Units","Profit_Margin_%","MoM_Growth_%"]].to_string(index=False))

print("\n--- Product Performance ---")
print(product.to_string(index=False))

print("\n--- Regional Performance ---")
print(region.to_string(index=False))

# ─────────────────────────────────────────
# 4. Dashboard Visualizations (6 charts)
# ─────────────────────────────────────────
os.makedirs("charts", exist_ok=True)
months = monthly["Month"]
x = range(len(months))

fig, axes = plt.subplots(3, 2, figsize=(18, 16))
fig.suptitle("Sales Analytics Dashboard — FY 2024", fontsize=18, fontweight="bold")

# Chart 1: Monthly Revenue vs Net Revenue
ax1 = axes[0, 0]
ax1.bar([i - 0.2 for i in x], monthly["Total_Revenue"], width=0.4, label="Gross Revenue", color="#2563eb")
ax1.bar([i + 0.2 for i in x], monthly["Net_Revenue"],   width=0.4, label="Net Revenue",   color="#16a34a")
ax1.set_title("Monthly Revenue vs Net Revenue", fontweight="bold")
ax1.set_xticks(list(x))
ax1.set_xticklabels(months, rotation=45, fontsize=8)
ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"₹{v/1e5:.0f}L"))
ax1.legend(); ax1.grid(axis="y", alpha=0.3)

# Chart 2: Sales Units Trend
ax2 = axes[0, 1]
ax2.plot(list(x), monthly["Total_Units"], marker="o", color="#7c3aed", linewidth=2.5, markersize=6)
ax2.fill_between(list(x), monthly["Total_Units"], alpha=0.15, color="#7c3aed")
ax2.set_title("Monthly Sales Units Trend", fontweight="bold")
ax2.set_xticks(list(x))
ax2.set_xticklabels(months, rotation=45, fontsize=8)
ax2.set_ylabel("Units Sold"); ax2.grid(alpha=0.3)

# Chart 3: Product-wise Revenue (Pie)
ax3 = axes[1, 0]
colors_pie = ["#2563eb", "#16a34a", "#f97316"]
wedges, texts, autotexts = ax3.pie(
    product["Total_Revenue"], labels=product["Product"],
    autopct="%1.1f%%", colors=colors_pie, startangle=90,
    wedgeprops={"edgecolor": "white", "linewidth": 2}
)
for at in autotexts: at.set_fontsize(10)
ax3.set_title("Product-wise Revenue Share", fontweight="bold")

# Chart 4: Region-wise Revenue (Bar)
ax4 = axes[1, 1]
region_colors = ["#2563eb", "#ef4444", "#16a34a"]
bars = ax4.bar(region["Region"], region["Total_Revenue"], color=region_colors, width=0.5)
ax4.set_title("Region-wise Revenue", fontweight="bold")
ax4.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"₹{v/1e6:.1f}M"))
ax4.grid(axis="y", alpha=0.3)
for bar, val in zip(bars, region["Total_Revenue"]):
    ax4.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 20000,
             f"₹{val/1e6:.1f}M", ha="center", fontsize=10, fontweight="bold")

# Chart 5: Profit Margin Trend
ax5 = axes[2, 0]
avg_margin = monthly["Profit_Margin_%"].mean()
bar_colors = ["#16a34a" if m >= avg_margin else "#f97316" for m in monthly["Profit_Margin_%"]]
bars5 = ax5.bar(list(x), monthly["Profit_Margin_%"], color=bar_colors)
ax5.axhline(y=avg_margin, color="red", linestyle="--", linewidth=1.5, label=f"Avg {avg_margin:.1f}%")
ax5.set_title("Monthly Profit Margin %", fontweight="bold")
ax5.set_xticks(list(x))
ax5.set_xticklabels(months, rotation=45, fontsize=8)
ax5.set_ylabel("Margin %"); ax5.legend(); ax5.grid(axis="y", alpha=0.3)

# Chart 6: MoM Revenue Growth
ax6 = axes[2, 1]
mom = monthly["MoM_Growth_%"].fillna(0)
colors_mom = ["#16a34a" if v >= 0 else "#ef4444" for v in mom]
ax6.bar(list(x), mom, color=colors_mom)
ax6.axhline(y=0, color="black", linewidth=0.8)
ax6.set_title("Month-over-Month Revenue Growth %", fontweight="bold")
ax6.set_xticks(list(x))
ax6.set_xticklabels(months, rotation=45, fontsize=8)
ax6.set_ylabel("Growth %"); ax6.grid(axis="y", alpha=0.3)

plt.tight_layout()
plt.savefig("charts/sales_dashboard.png", dpi=150, bbox_inches="tight")
plt.close()
print("\n✅ Dashboard saved: charts/sales_dashboard.png")

# ─────────────────────────────────────────
# 5. Export Excel Report
# ─────────────────────────────────────────
with pd.ExcelWriter("Sales_Analytics_Report.xlsx", engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Raw Data", index=False)
    monthly.to_excel(writer, sheet_name="Monthly Trend", index=False)
    product.to_excel(writer, sheet_name="Product Analysis", index=False)
    region.to_excel(writer, sheet_name="Region Analysis", index=False)

    kpi_df = pd.DataFrame({
        "KPI": ["Total Revenue", "Total Net Revenue", "Total Units Sold",
                "Total Returns", "Average ROI %", "Best Product", "Best Region", "Best Month"],
        "Value": [f"₹{total_rev:,.0f}", f"₹{total_net:,.0f}", f"{total_units:,}",
                  f"{total_ret:,}", f"{avg_roi:.2f}%", best_product, best_region, best_month]
    })
    kpi_df.to_excel(writer, sheet_name="KPI Summary", index=False)

print("✅ Excel report saved: Sales_Analytics_Report.xlsx")
print("\n✅ Analysis Complete!")
