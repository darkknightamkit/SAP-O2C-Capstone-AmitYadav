"""
SAP Order-to-Cash (O2C) — Data Cleaning & Exploratory Data Analysis
Student: Amit Yadav | Roll No: 23052297 | Batch: CSE33 — Data Analytics
"""

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import seaborn as sns
import warnings, os
warnings.filterwarnings('ignore')

# ── Paths ─────────────────────────────────────────────────
DATA_PATH    = os.path.join(os.path.dirname(__file__), '..', 'data', 'sap_o2c_sales_data.csv')
REPORTS_PATH = os.path.join(os.path.dirname(__file__), '..', 'reports')
os.makedirs(REPORTS_PATH, exist_ok=True)

sns.set_theme(style='whitegrid', palette='muted')
COLORS = ['#1565c0','#42a5f5','#ef5350','#66bb6a','#ffa726','#ab47bc','#26c6da','#8d6e63']

# ════════════════════════════════════════════════════════
# 1. LOAD & INSPECT
# ════════════════════════════════════════════════════════
print("=" * 60)
print("  SAP O2C Data Analysis — Amit Yadav (23052297)")
print("=" * 60)

df = pd.read_csv(DATA_PATH)
print(f"\n[1] Dataset Shape  : {df.shape[0]} rows × {df.shape[1]} columns")
print(f"    Columns        : {list(df.columns)}")
print(f"    Date Range     : {df['Order_Date'].min()} → {df['Order_Date'].max()}")

# ════════════════════════════════════════════════════════
# 2. DATA CLEANING
# ════════════════════════════════════════════════════════
print("\n[2] Data Cleaning")

# Parse dates
for col in ['Order_Date','Delivery_Date','Invoice_Date']:
    df[col] = pd.to_datetime(df[col])

# Fill missing Payment_Days with median for paid orders
median_pay = df[df['Payment_Days'].notna()]['Payment_Days'].median()
df['Payment_Days'] = df['Payment_Days'].fillna(0)

# Remove cancelled orders for revenue analysis (keep for status analysis)
df_clean = df[df['Order_Status'] != 'Cancelled'].copy()

missing = df.isnull().sum()
print(f"    Missing values before clean : {missing[missing > 0].to_dict()}")
print(f"    Records after removing cancelled: {len(df_clean)}")
print(f"    Duplicate rows : {df.duplicated().sum()}")

# ════════════════════════════════════════════════════════
# 3. KPI SUMMARY
# ════════════════════════════════════════════════════════
print("\n[3] Key Performance Indicators")
total_revenue  = df_clean['Net_Value_INR'].sum()
total_profit   = df_clean['Gross_Profit_INR'].sum()
total_orders   = len(df_clean)
avg_order_val  = df_clean['Net_Value_INR'].mean()
avg_margin     = df_clean['Profit_Margin_Pct'].mean()
avg_del_days   = df_clean['Delivery_Days'].mean()

print(f"    Total Revenue      : ₹{total_revenue:,.2f}")
print(f"    Total Gross Profit : ₹{total_profit:,.2f}")
print(f"    Total Orders       : {total_orders}")
print(f"    Avg Order Value    : ₹{avg_order_val:,.2f}")
print(f"    Avg Profit Margin  : {avg_margin:.1f}%")
print(f"    Avg Delivery Days  : {avg_del_days:.1f}")

# ════════════════════════════════════════════════════════
# 4. CHARTS
# ════════════════════════════════════════════════════════

# ── Chart 1: Monthly Revenue Trend ───────────────────────
monthly = df_clean.groupby(['Year','Month'])['Net_Value_INR'].sum().reset_index()
month_order = ['January','February','March','April','May','June',
               'July','August','September','October','November','December']
monthly['Month'] = pd.Categorical(monthly['Month'], categories=month_order, ordered=True)
monthly = monthly.sort_values(['Year','Month'])
monthly['Label'] = monthly['Month'].astype(str).str[:3]

fig, ax = plt.subplots(figsize=(12, 5))
ax.fill_between(range(len(monthly)), monthly['Net_Value_INR']/1e5,
                alpha=0.25, color='#1565c0')
ax.plot(range(len(monthly)), monthly['Net_Value_INR']/1e5,
        marker='o', color='#1565c0', linewidth=2.5, markersize=6)
ax.set_xticks(range(len(monthly)))
ax.set_xticklabels(monthly['Label'], rotation=45, fontsize=9)
ax.set_title('Monthly Net Revenue Trend (2024)', fontsize=14, fontweight='bold', pad=12)
ax.set_ylabel('Revenue (₹ Lakhs)', fontsize=11)
ax.set_xlabel('Month', fontsize=11)
fig.tight_layout()
fig.savefig(f'{REPORTS_PATH}/01_monthly_revenue.png', dpi=150)
plt.close()
print("\n    [Chart 1] Monthly Revenue saved.")

# ── Chart 2: Revenue by Customer ─────────────────────────
cust_rev = df_clean.groupby('Customer_Name')['Net_Value_INR'].sum().nlargest(8).reset_index()
fig, ax = plt.subplots(figsize=(10, 5))
bars = ax.barh(cust_rev['Customer_Name'], cust_rev['Net_Value_INR']/1e5,
               color=COLORS[:len(cust_rev)])
ax.set_xlabel('Revenue (₹ Lakhs)', fontsize=11)
ax.set_title('Top 8 Customers by Revenue', fontsize=14, fontweight='bold', pad=12)
for bar, val in zip(bars, cust_rev['Net_Value_INR']/1e5):
    ax.text(bar.get_width() + 0.5, bar.get_y() + bar.get_height()/2,
            f'₹{val:.1f}L', va='center', fontsize=9)
fig.tight_layout()
fig.savefig(f'{REPORTS_PATH}/02_top_customers.png', dpi=150)
plt.close()
print("    [Chart 2] Top Customers saved.")

# ── Chart 3: Category Revenue Pie ────────────────────────
cat_rev = df_clean.groupby('Category')['Net_Value_INR'].sum()
fig, ax = plt.subplots(figsize=(7, 7))
wedges, texts, autotexts = ax.pie(
    cat_rev.values, labels=cat_rev.index,
    autopct='%1.1f%%', colors=COLORS[:len(cat_rev)],
    startangle=140, pctdistance=0.82,
    wedgeprops=dict(edgecolor='white', linewidth=2))
for t in autotexts: t.set_fontsize(11)
ax.set_title('Revenue Share by Product Category', fontsize=14, fontweight='bold')
fig.tight_layout()
fig.savefig(f'{REPORTS_PATH}/03_category_revenue.png', dpi=150)
plt.close()
print("    [Chart 3] Category Pie saved.")

# ── Chart 4: Order Status Distribution ───────────────────
status = df['Order_Status'].value_counts()
fig, ax = plt.subplots(figsize=(8, 5))
bars = ax.bar(status.index, status.values, color=COLORS[:len(status)], edgecolor='white', linewidth=1.5)
ax.set_title('Order Status Distribution', fontsize=14, fontweight='bold', pad=12)
ax.set_ylabel('Number of Orders', fontsize=11)
for bar, val in zip(bars, status.values):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 2,
            str(val), ha='center', fontsize=11, fontweight='bold')
fig.tight_layout()
fig.savefig(f'{REPORTS_PATH}/04_order_status.png', dpi=150)
plt.close()
print("    [Chart 4] Order Status saved.")

# ── Chart 5: Sales Rep Performance ───────────────────────
rep_perf = df_clean.groupby('Sales_Rep').agg(
    Revenue=('Net_Value_INR','sum'),
    Orders=('Sales_Order_No','count'),
    Margin=('Profit_Margin_Pct','mean')
).reset_index()
rep_perf['Rep'] = rep_perf['Sales_Rep'].str.split('-').str[1]

fig, axes = plt.subplots(1, 2, figsize=(12, 5))
axes[0].bar(rep_perf['Rep'], rep_perf['Revenue']/1e5, color=COLORS[:5], edgecolor='white')
axes[0].set_title('Revenue by Sales Rep (₹ Lakhs)', fontsize=12, fontweight='bold')
axes[0].tick_params(axis='x', rotation=20)

axes[1].bar(rep_perf['Rep'], rep_perf['Margin'], color=COLORS[2:7], edgecolor='white')
axes[1].set_title('Avg Profit Margin % by Sales Rep', fontsize=12, fontweight='bold')
axes[1].set_ylabel('%')
axes[1].tick_params(axis='x', rotation=20)

fig.tight_layout()
fig.savefig(f'{REPORTS_PATH}/05_sales_rep_performance.png', dpi=150)
plt.close()
print("    [Chart 5] Sales Rep Performance saved.")

# ── Chart 6: Domestic vs Export ──────────────────────────
ctype = df_clean.groupby('Customer_Type')['Net_Value_INR'].sum()
fig, ax = plt.subplots(figsize=(6, 5))
ax.bar(ctype.index, ctype.values/1e5, color=['#1565c0','#ef5350'], edgecolor='white', linewidth=2)
ax.set_title('Domestic vs Export Revenue', fontsize=13, fontweight='bold')
ax.set_ylabel('Revenue (₹ Lakhs)')
for i, val in enumerate(ctype.values/1e5):
    ax.text(i, val + 0.5, f'₹{val:.1f}L', ha='center', fontweight='bold')
fig.tight_layout()
fig.savefig(f'{REPORTS_PATH}/06_domestic_vs_export.png', dpi=150)
plt.close()
print("    [Chart 6] Domestic vs Export saved.")

# ── Chart 7: Quarterly Revenue ────────────────────────────
qtr = df_clean.groupby('Quarter')['Net_Value_INR'].sum().reset_index()
fig, ax = plt.subplots(figsize=(7, 5))
ax.bar(qtr['Quarter'], qtr['Net_Value_INR']/1e5, color=COLORS[:4], edgecolor='white', linewidth=2, width=0.5)
ax.set_title('Quarterly Revenue (2024)', fontsize=13, fontweight='bold')
ax.set_ylabel('Revenue (₹ Lakhs)')
for i, row in qtr.iterrows():
    ax.text(i, row['Net_Value_INR']/1e5 + 0.5, f"₹{row['Net_Value_INR']/1e5:.1f}L",
            ha='center', fontweight='bold')
fig.tight_layout()
fig.savefig(f'{REPORTS_PATH}/07_quarterly_revenue.png', dpi=150)
plt.close()
print("    [Chart 7] Quarterly Revenue saved.")

# ── Chart 8: Profit Margin by Category ───────────────────
cat_margin = df_clean.groupby('Category')['Profit_Margin_Pct'].mean().sort_values()
fig, ax = plt.subplots(figsize=(8, 4))
ax.barh(cat_margin.index, cat_margin.values, color=COLORS[:3], edgecolor='white')
ax.set_title('Average Profit Margin by Category (%)', fontsize=13, fontweight='bold')
ax.set_xlabel('Profit Margin %')
for i, v in enumerate(cat_margin.values):
    ax.text(v + 0.2, i, f'{v:.1f}%', va='center', fontweight='bold')
fig.tight_layout()
fig.savefig(f'{REPORTS_PATH}/08_margin_by_category.png', dpi=150)
plt.close()
print("    [Chart 8] Margin by Category saved.")

# ════════════════════════════════════════════════════════
# 5. EXPORT SUMMARY EXCEL REPORT
# ════════════════════════════════════════════════════════
print("\n[4] Exporting Summary Excel Report...")
with pd.ExcelWriter(f'{REPORTS_PATH}/SAP_O2C_Summary_Report.xlsx', engine='openpyxl') as writer:
    # KPI sheet
    kpi_df = pd.DataFrame({
        'KPI': ['Total Revenue (INR)', 'Total Gross Profit (INR)', 'Total Orders',
                'Avg Order Value (INR)', 'Avg Profit Margin (%)', 'Avg Delivery Days'],
        'Value': [f'₹{total_revenue:,.2f}', f'₹{total_profit:,.2f}', total_orders,
                  f'₹{avg_order_val:,.2f}', f'{avg_margin:.1f}%', f'{avg_del_days:.1f}']
    })
    kpi_df.to_excel(writer, sheet_name='KPI Summary', index=False)

    # Monthly
    monthly_out = df_clean.groupby(['Month','Quarter']).agg(
        Revenue=('Net_Value_INR','sum'),
        Orders=('Sales_Order_No','count'),
        Profit=('Gross_Profit_INR','sum')
    ).reset_index()
    monthly_out.to_excel(writer, sheet_name='Monthly Revenue', index=False)

    # Customer
    cust_out = df_clean.groupby(['Customer_ID','Customer_Name','Customer_Type']).agg(
        Revenue=('Net_Value_INR','sum'),
        Orders=('Sales_Order_No','count'),
        Margin=('Profit_Margin_Pct','mean')
    ).reset_index().sort_values('Revenue', ascending=False)
    cust_out.to_excel(writer, sheet_name='Customer Analysis', index=False)

    # Product
    prod_out = df_clean.groupby(['Material_Code','Category']).agg(
        Revenue=('Net_Value_INR','sum'),
        Qty=('Quantity','sum'),
        Margin=('Profit_Margin_Pct','mean')
    ).reset_index().sort_values('Revenue', ascending=False)
    prod_out.to_excel(writer, sheet_name='Product Analysis', index=False)

    # Full data
    df_clean.to_excel(writer, sheet_name='Raw Data', index=False)

print(f"    Excel report saved → {REPORTS_PATH}/SAP_O2C_Summary_Report.xlsx")
print("\n✅ All analysis complete!")
