"""
Order to Cash – Cash Application & Invoice Reconciliation System
Author      : Narmadhadevi C
Dataset     : Real-world O2C Invoice Dataset (Kaggle) — 50,000 records
Description : Applies payments to open invoices, performs reconciliation,
              aging analysis, and generates a complete O2C Excel report.
"""

import pandas as pd
import numpy as np
from datetime import datetime
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import warnings
warnings.filterwarnings("ignore")

print("=" * 60)
print("   ORDER TO CASH — CASH APPLICATION & RECONCILIATION")
print("=" * 60)

# ─────────────────────────────────────────────────────
# 1. LOAD & CLEAN DATA
# ─────────────────────────────────────────────────────
df = pd.read_csv("dataset.csv")

print(f"\n✅ Dataset Loaded: {len(df):,} records")

# Clean & parse dates
def safe_date(val):
    try:
        s = str(int(float(val)))
        return pd.to_datetime(s, format="%Y%m%d", errors="coerce")
    except:
        return pd.NaT

df["due_date"]      = df["due_in_date"].apply(safe_date)
df["create_date"]   = df["document_create_date"].apply(safe_date)
df["clear_date"]    = pd.to_datetime(df["clear_date"], errors="coerce")
df["posting_date"]  = pd.to_datetime(df["posting_date"], errors="coerce")
df["amount"]        = pd.to_numeric(df["total_open_amount"], errors="coerce").fillna(0)
df["name_customer"] = df["name_customer"].fillna("Unknown Customer").str.strip()

today = pd.Timestamp("2020-06-30")

# ─────────────────────────────────────────────────────
# 2. CASH APPLICATION — Paid vs Open
# ─────────────────────────────────────────────────────
paid_df = df[df["isOpen"] == 0].copy()
open_df = df[df["isOpen"] == 1].copy()

print(f"\n📋 Invoice Summary")
print(f"   Total Invoices     : {len(df):,}")
print(f"   Paid / Cleared     : {len(paid_df):,}")
print(f"   Open / Outstanding : {len(open_df):,}")

# Cash Application Status for paid invoices
def cash_app_status(row):
    if pd.isna(row["clear_date"]) or pd.isna(row["due_date"]):
        return "Cleared – Date Unknown"
    if row["clear_date"] <= row["due_date"]:
        return "Paid – On Time"
    else:
        days_late = (row["clear_date"] - row["due_date"]).days
        if days_late <= 15:
            return "Paid – Slightly Late (≤15 days)"
        elif days_late <= 30:
            return "Paid – Late (16–30 days)"
        else:
            return "Paid – Very Late (>30 days)"

paid_df["application_status"] = paid_df.apply(cash_app_status, axis=1)
paid_df["days_to_pay"] = (paid_df["clear_date"] - paid_df["due_date"]).dt.days

# ─────────────────────────────────────────────────────
# 3. AGING ANALYSIS — Open Invoices
# ─────────────────────────────────────────────────────
def aging_bucket(row):
    if pd.isna(row["due_date"]):
        return "Unknown"
    days = (today - row["due_date"]).days
    if days <= 0:
        return "Current (Not Yet Due)"
    elif days <= 30:
        return "1–30 Days Overdue"
    elif days <= 60:
        return "31–60 Days Overdue"
    elif days <= 90:
        return "61–90 Days Overdue"
    else:
        return "90+ Days Overdue"

open_df["aging_bucket"] = open_df.apply(aging_bucket, axis=1)
open_df["days_overdue"] = (today - open_df["due_date"]).dt.days.clip(lower=0)

# ─────────────────────────────────────────────────────
# 4. SUMMARY STATS
# ─────────────────────────────────────────────────────
total_invoiced  = df["amount"].sum()
total_collected = paid_df["amount"].sum()
total_open      = open_df["amount"].sum()
collection_rate = (total_collected / total_invoiced * 100) if total_invoiced > 0 else 0
avg_days_late   = paid_df["days_to_pay"].mean()

print(f"\n💰 Financial Summary")
print(f"   Total Invoiced     : ${total_invoiced:>15,.2f}")
print(f"   Total Collected    : ${total_collected:>15,.2f}")
print(f"   Total Outstanding  : ${total_open:>15,.2f}")
print(f"   Collection Rate    :  {collection_rate:>14.1f}%")
print(f"   Avg Days to Pay    :  {avg_days_late:>14.1f} days")

print(f"\n📊 Cash Application Status Breakdown:")
print(paid_df["application_status"].value_counts().to_string())

print(f"\n📅 AR Aging Breakdown (Open Invoices):")
aging_summary = open_df.groupby("aging_bucket")["amount"].agg(["count","sum"])
aging_summary.columns = ["Count", "Amount ($)"]
aging_summary["Amount ($)"] = aging_summary["Amount ($)"].map("${:,.2f}".format)
print(aging_summary.to_string())

# ─────────────────────────────────────────────────────
# 5. TOP CUSTOMERS — Outstanding
# ─────────────────────────────────────────────────────
top_customers = (
    open_df.groupby("name_customer")["amount"]
    .sum().sort_values(ascending=False).head(10).reset_index()
)
top_customers.columns = ["Customer Name", "Outstanding Amount ($)"]

print(f"\n🏆 Top 10 Customers by Outstanding Amount:")
print(top_customers.to_string(index=False))

# ─────────────────────────────────────────────────────
# 6. PREPARE EXPORT DATA
# ─────────────────────────────────────────────────────
cash_app_export = paid_df[[
    "invoice_id","name_customer","business_code","invoice_currency",
    "amount","create_date","due_date","clear_date",
    "days_to_pay","application_status"
]].copy().head(5000)
cash_app_export.columns = [
    "Invoice ID","Customer Name","Business Code","Currency",
    "Invoice Amount ($)","Invoice Date","Due Date","Payment Date",
    "Days to Pay","Application Status"
]
for col in ["Invoice Date","Due Date","Payment Date"]:
    cash_app_export[col] = pd.to_datetime(cash_app_export[col]).dt.strftime("%d-%b-%Y")

aging_export = open_df[[
    "invoice_id","name_customer","business_code","invoice_currency",
    "amount","due_date","days_overdue","aging_bucket"
]].copy()
aging_export.columns = [
    "Invoice ID","Customer Name","Business Code","Currency",
    "Outstanding Amount ($)","Due Date","Days Overdue","Aging Bucket"
]
aging_export["Due Date"] = pd.to_datetime(aging_export["Due Date"]).dt.strftime("%d-%b-%Y")

aging_pivot = open_df.groupby("aging_bucket")["amount"].agg(
    Count="count", Total_Amount="sum"
).reset_index()
aging_pivot.columns = ["Aging Bucket","Invoice Count","Total Outstanding ($)"]

top_cust_export = top_customers.copy()

summary_df = pd.DataFrame({
    "Metric": [
        "Total Invoices","Paid Invoices","Open Invoices",
        "Total Invoiced ($)","Total Collected ($)",
        "Total Outstanding ($)","Collection Rate (%)","Avg Days to Pay"
    ],
    "Value": [
        f"{len(df):,}", f"{len(paid_df):,}", f"{len(open_df):,}",
        f"${total_invoiced:,.2f}", f"${total_collected:,.2f}",
        f"${total_open:,.2f}", f"{collection_rate:.1f}%",
        f"{avg_days_late:.1f} days"
    ]
})

# ─────────────────────────────────────────────────────
# 7. EXPORT TO EXCEL
# ─────────────────────────────────────────────────────
output_file = "O2C_Full_Report.xlsx"
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    summary_df.to_excel(writer,       sheet_name="Executive Summary",  index=False)
    cash_app_export.to_excel(writer,  sheet_name="Cash Application",   index=False)
    aging_export.to_excel(writer,     sheet_name="AR Aging Detail",    index=False)
    aging_pivot.to_excel(writer,      sheet_name="Aging Buckets",      index=False)
    top_cust_export.to_excel(writer,  sheet_name="Top 10 Outstanding", index=False)

# ─────────────────────────────────────────────────────
# 8. FORMAT EXCEL
# ─────────────────────────────────────────────────────
STATUS_COLORS = {
    "Paid – On Time"                  : "C6EFCE",
    "Paid – Slightly Late (≤15 days)" : "FFEB9C",
    "Paid – Late (16–30 days)"        : "FFCC99",
    "Paid – Very Late (>30 days)"     : "FFC7CE",
    "Cleared – Date Unknown"          : "DDEBF7",
    "Current (Not Yet Due)"           : "C6EFCE",
    "1–30 Days Overdue"               : "FFEB9C",
    "31–60 Days Overdue"              : "FFCC99",
    "61–90 Days Overdue"              : "FFC7CE",
    "90+ Days Overdue"                : "FF4444",
}

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
CENTER      = Alignment(horizontal="center", vertical="center", wrap_text=True)
thin        = Side(style="thin", color="CCCCCC")
BORDER      = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = openpyxl.load_workbook(output_file)

for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]

    for cell in ws[1]:
        cell.fill      = HEADER_FILL
        cell.font      = HEADER_FONT
        cell.alignment = CENTER
        cell.border    = BORDER

    color_col = None
    for idx, cell in enumerate(ws[1], 1):
        if cell.value in ("Application Status", "Aging Bucket"):
            color_col = idx

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border    = BORDER
            cell.alignment = Alignment(vertical="center")
        if color_col:
            val   = row[color_col - 1].value
            color = STATUS_COLORS.get(str(val), "FFFFFF")
            if color != "FFFFFF":
                for cell in row:
                    cell.fill = PatternFill("solid", fgColor=color)

    for col in ws.columns:
        max_len = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 30)

    ws.freeze_panes = "A2"

wb.save(output_file)
print(f"\n✅ Full O2C Report saved: {output_file}")
print("🎉 All done! 5 sheets generated successfully.")
print("=" * 60)
