 SAP Order-to-Cash (O2C) — Capstone Project



---

 📌 Project Overview

This capstone project implements a complete **SAP Order-to-Cash (O2C)** analytics solution
covering the full sales cycle from Sales Order creation to Payment Receipt — the core
process in SAP SD (Sales & Distribution) module.

The project includes:
- Realistic SAP-style transactional dataset (500 sales orders)
- Python-based data cleaning and EDA pipeline
- 8 analytical charts (saved as PNG)
- Multi-sheet Excel summary report
- Interactive HTML dashboard with Chart.js

---

 📁 Folder Structure

```
SAP_O2C_Project/
├── data/
│   └── sap_o2c_sales_data.csv       ← 500 SAP O2C transactions
├── scripts/
│   ├── generate_data.py             ← Dataset generation script
│   └── data_analysis.py             ← EDA, charts, Excel export
├── reports/
│   ├── 01_monthly_revenue.png
│   ├── 02_top_customers.png
│   ├── 03_category_revenue.png
│   ├── 04_order_status.png
│   ├── 05_sales_rep_performance.png
│   ├── 06_domestic_vs_export.png
│   ├── 07_quarterly_revenue.png
│   ├── 08_margin_by_category.png
│   └── SAP_O2C_Summary_Report.xlsx  ← Multi-sheet summary
├── dashboard/
│   └── SAP_O2C_Dashboard.html       ← Interactive dashboard (open in browser)
├── docs/
│   └── Amit_Yadav_23052297_Capstone_Project.pdf ← Project documentation PDF
└── README.md
```

---

 🔄 SAP O2C Process Flow

```
[Inquiry / Quotation]
        ↓
[Sales Order (VA01)]       ← Doc Types: OR, ZOR, ZEX
        ↓
[Delivery Document (VL01N)]
        ↓
[Goods Issue (PGI)]
        ↓
[Billing / Invoice (VF01)]
        ↓
[Payment Receipt (F-28)]
```

---

⚙️ Tech Stack

| Component         | Technology              |
|-------------------|------------------------|
| Programming Lang  | Python 3.11             |
| Data Processing   | Pandas, NumPy           |
| Visualization     | Matplotlib, Seaborn     |
| Report Export     | OpenPyXL (Excel)        |
| Dashboard         | HTML5 + Chart.js        |
| Dataset Format    | CSV                     |
| IDE               | Jupyter / VS Code       |

---

## 🚀 How to Run

### Step 1 — Install Dependencies
```bash
pip install pandas numpy matplotlib seaborn openpyxl
```

### Step 2 — Generate Dataset
```bash
python scripts/generate_data.py
```

### Step 3 — Run Analysis (Charts + Excel)
```bash
python scripts/data_analysis.py
```

### Step 4 — View Dashboard
Open `dashboard/SAP_O2C_Dashboard.html` in any web browser.

---

 📊 Key Results

| KPI                    | Value           |
|------------------------|-----------------|
| Total Revenue          | ₹16.30 Crore    |
| Gross Profit           | ₹6.19 Crore     |
| Avg Profit Margin      | 37.7%           |
| Total Orders (Active)  | 481             |
| Avg Delivery Days      | 12.4 days       |
| Payment Received       | 285 orders (59%)|
| Export Revenue Share   | ~28%            |

---

 📋 SAP Modules Covered

- SAP SD — Sales Order, Delivery, Billing
- SAP FI — Invoice posting, Payment clearing
- SAP MM — Material master, Plant data
- SAP CO — Profitability analysis (COPA)

---


