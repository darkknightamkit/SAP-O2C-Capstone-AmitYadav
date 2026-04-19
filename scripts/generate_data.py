"""
SAP Order-to-Cash (O2C) - Data Generator
Generates realistic SAP-style sales transaction data
Student: Amit Yadav | Roll: 23052297 | Batch: CSE33
"""

import pandas as pd
import numpy as np
import os
import random
from datetime import datetime, timedelta

np.random.seed(42)
random.seed(42)

# ── Master Data ──────────────────────────────────────────
customers = {
    'C001': ('Reliance Industries Ltd',    'Mumbai',    'Maharashtra', 'IN', 'Domestic'),
    'C002': ('Tata Consultancy Services',  'Pune',      'Maharashtra', 'IN', 'Domestic'),
    'C003': ('Infosys Technologies',       'Bangalore', 'Karnataka',   'IN', 'Domestic'),
    'C004': ('Wipro Ltd',                  'Hyderabad', 'Telangana',   'IN', 'Domestic'),
    'C005': ('HCL Technologies',           'Noida',     'Uttar Pradesh','IN','Domestic'),
    'C006': ('Mahindra & Mahindra',        'Chennai',   'Tamil Nadu',  'IN', 'Domestic'),
    'C007': ('Bajaj Auto Ltd',             'Pune',      'Maharashtra', 'IN', 'Domestic'),
    'C008': ('Sun Pharma',                 'Ahmedabad', 'Gujarat',     'IN', 'Domestic'),
    'C009': ('HDFC Bank',                  'Delhi',     'Delhi',       'IN', 'Domestic'),
    'C010': ('ITC Limited',               'Kolkata',   'West Bengal', 'IN', 'Domestic'),
    'C011': ('Global Tech Corp',           'New York',  'NY',          'US', 'Export'),
    'C012': ('Euro Solutions GmbH',        'Berlin',    'Berlin',      'DE', 'Export'),
    'C013': ('Asia Pacific Traders',       'Singapore', 'SG',          'SG', 'Export'),
}

products = {
    'MAT-1001': ('Industrial Pump A100',   'Mechanical', 15000, 9000),
    'MAT-1002': ('Hydraulic Valve V200',   'Mechanical', 8500,  5100),
    'MAT-1003': ('Control Panel CP300',    'Electrical', 22000, 13200),
    'MAT-1004': ('Pressure Sensor PS400',  'Electrical', 3500,  2100),
    'MAT-1005': ('Steel Pipe Grade A',     'Raw Material',1200, 720),
    'MAT-1006': ('Electric Motor EM600',   'Electrical', 18000, 10800),
    'MAT-1007': ('Gear Box GB700',         'Mechanical', 12000, 7200),
    'MAT-1008': ('Safety Valve SV800',     'Mechanical', 6000,  3600),
    'MAT-1009': ('PLC Unit PL900',         'Electrical', 35000, 21000),
    'MAT-1010': ('Conveyor Belt CB1000',   'Raw Material',9000, 5400),
}

sales_reps = ['SR001-Rahul Sharma', 'SR002-Priya Patel',
              'SR003-Arjun Singh',  'SR004-Meena Nair', 'SR005-Vikram Rao']

doc_types   = ['OR', 'OR', 'OR', 'ZOR', 'ZEX']   # Standard/Custom/Export orders
divisions   = ['10', '20', '30']
plants      = ['PL01-Mumbai', 'PL02-Pune', 'PL03-Bangalore']
ship_points = ['SP01', 'SP02', 'SP03']

# ── Generate Sales Orders ─────────────────────────────────
start_date = datetime(2024, 1, 1)
records = []

for i in range(500):
    cust_id  = random.choice(list(customers.keys()))
    cust_info = customers[cust_id]
    mat_id   = random.choice(list(products.keys()))
    mat_info = products[mat_id]

    order_date    = start_date + timedelta(days=random.randint(0, 364))
    delivery_days = random.randint(3, 21)
    delivery_date = order_date + timedelta(days=delivery_days)
    invoice_date  = delivery_date + timedelta(days=random.randint(1, 5))
    payment_days  = random.randint(0, 60)
    payment_date  = invoice_date + timedelta(days=payment_days)

    qty       = random.randint(1, 50)
    unit_price = mat_info[2] * random.uniform(0.95, 1.10)
    discount   = random.uniform(0, 0.12)
    net_value  = qty * unit_price * (1 - discount)
    cogs       = qty * mat_info[3]
    profit     = net_value - cogs

    doc_type  = 'ZEX' if cust_info[4] == 'Export' else random.choice(['OR','OR','ZOR'])
    status    = random.choices(
        ['Delivered', 'Invoiced', 'Payment Received', 'Open', 'Cancelled'],
        weights=[10, 15, 60, 10, 5])[0]

    records.append({
        'Sales_Order_No':   f'SO-{2024000 + i + 1}',
        'Doc_Type':         doc_type,
        'Order_Date':       order_date.strftime('%Y-%m-%d'),
        'Delivery_Date':    delivery_date.strftime('%Y-%m-%d'),
        'Invoice_Date':     invoice_date.strftime('%Y-%m-%d'),
        'Payment_Date':     payment_date.strftime('%Y-%m-%d') if status == 'Payment Received' else '',
        'Customer_ID':      cust_id,
        'Customer_Name':    cust_info[0],
        'City':             cust_info[1],
        'State':            cust_info[2],
        'Country':          cust_info[3],
        'Customer_Type':    cust_info[4],
        'Material_Code':    mat_id,
        'Material_Desc':    mat_info[1],
        'Category':         mat_info[1],
        'Plant':            random.choice(plants),
        'Shipping_Point':   random.choice(ship_points),
        'Division':         random.choice(divisions),
        'Sales_Rep':        random.choice(sales_reps),
        'Quantity':         qty,
        'Unit_Price_INR':   round(unit_price, 2),
        'Discount_Pct':     round(discount * 100, 2),
        'Net_Value_INR':    round(net_value, 2),
        'COGS_INR':         round(cogs, 2),
        'Gross_Profit_INR': round(profit, 2),
        'Profit_Margin_Pct':round((profit / net_value) * 100, 2) if net_value else 0,
        'Delivery_Days':    delivery_days,
        'Payment_Days':     payment_days if status == 'Payment Received' else None,
        'Order_Status':     status,
        'Month':            order_date.strftime('%B'),
        'Quarter':          f"Q{(order_date.month - 1) // 3 + 1}",
        'Year':             order_date.year,
    })

df = pd.DataFrame(records)
out = '/home/claude/SAP_O2C_Project/data/sap_o2c_sales_data.csv'
df.to_csv(out, index=False)
print(f"Dataset generated: {len(df)} records → {out}")
print(df.head(3).to_string())
