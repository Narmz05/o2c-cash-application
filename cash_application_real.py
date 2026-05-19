"""
Order to Cash – Cash Application & Invoice Reconciliation System

Author      : Narmadhadevi C
Dataset     : Real-world O2C Invoice Dataset (Kaggle) — 50,000 records
Description : Applies payments to open invoices, performs reconciliation,
              aging analysis, and generates a complete O2C Excel report.

Fixes applied:
  FIX 1 – Aging uses dataset's max posting_date as reference, not live today
  FIX 2 – days_to_pay split into early vs late; avg only over late invoices
  FIX 3 – "Cleared - Date Unknown" replaced with "Data Quality Issue" label
  FIX 4 – Zero-amount invoices excluded from financial KPI calculations
  FIX 5 – "Unknown Customer" excluded from Top 10 outstanding list
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
print("  ORDER TO CASH — CASH APPLICATION & RECONCILIATION")
print("=" * 60)

# ─────────────────────────────────────────────
# 1. LOAD & CLEAN DATA
# ─────────────────────────────────────────────
df = pd.read_csv("dataset.csv")
print(f"\nDataset Loaded: {len(df):,} records")


def safe_date(val):
    """Parse SAP-style integer dates (YYYYMMDD) safely."""
    try:
        s = str(int(float(val)))
        return pd.to_datetime(s, format="%Y%m%d", errors="coerce")
    except Exception:
        return pd.NaT


df["due_date"]     = df["due_in_date"].apply(safe_date)
df["create_date"]  = df["document_create_date"].apply(safe_date)
df["clear_date"]   = pd.to_datetime(df["clear_date"], errors="coerce")
df["posting_date"] = pd.to_datetime(df["posting_date"], errors="coerce")
df["amount"]       = pd.to_numeric(df["total_open_amount"], errors="coerce").fillna(0)
df["name_customer"] = df["name_customer"].fillna("Unknown Customer").str.strip()

# FIX 1 — Use dataset's own max posting_date as the "as-of" reference date.
#          Using pd.Timestamp.today() on a historical dataset makes every open
#          invoice fall into "90+ Days Overdue", rendering aging meaningless.
AS_OF_DATE = df["posting_date"].max()
print(f"\nAging reference date (as-of): {AS_OF_DATE.strftime('%d-%b-%Y')}")

# FIX 4 — Exclude zero-amount rows (credit memos, reversals) from financials.
#          These inflate invoice counts and distort collection rate.
df_financial = df[df["amount"] != 0].copy()

# ─────────────────────────────────────────────
# 2. CASH APPLICATION
# ─────────────────────────────────────────────
paid_df = df_financial[df_financial["isOpen"] == 0].copy()
open_df = df_financial[df_financial["isOpen"] == 1].copy()

# Also keep zero-amount rows for count reporting (not financials)
total_all = len(df)
total_paid_all = len(df[df["isOpen"] == 0])
total_open_all = len(df[df["isOpen"] == 1])

print(f"\nInvoice Summary (all rows)")
print(f"  Total Invoices       : {total_all:,}")
print(f"  Paid / Cleared       : {total_paid_all:,}")
print(f"  Open / Outstanding   : {total_open_all:,}")
print(f"\nFinancial rows (amount ≠ 0): {len(df_financial):,}")


def cash_app_status(row):
    """
    Classify each paid invoice by payment timeliness.

    FIX 3 — If clear_date is missing on an isOpen=0 record, it is a data
             quality issue, NOT a confirmed clearance. Labelling it 'Cleared'
             was misleading; it is now explicitly flagged.
    """
    if pd.isna(row["due_date"]):
        return "Data Quality Issue – No Due Date"
    if pd.isna(row["clear_date"]):
        return "Data Quality Issue – No Clear Date"

    if row["clear_date"] <= row["due_date"]:
        return "Paid – On Time"

    days_late = (row["clear_date"] - row["due_date"]).days
    if days_late <= 15:
        return "Paid – Slightly Late (≤15 days)"
    elif days_late <= 30:
        return "Paid – Late (16–30 days)"
    else:
        return "Paid – Very Late (>30 days)"


paid_df["application_status"] = paid_df.apply(cash_app_status, axis=1)

# FIX 2 — Raw days_to_pay can be negative (early payments), which drags down
#          the average unfairly. We track both directions separately.
paid_df["days_to_pay"] = (paid_df["clear_date"] - paid_df["due_date"]).dt.days

# days_early: how early a payment was (positive = paid before due date)
paid_df["days_early"] = (-paid_df["days_to_pay"]).clip(lower=0)
# days_late:  how many days past due (positive = overdue at payment)
paid_df["days_late"]  = paid_df["days_to_pay"].clip(lower=0)

# ─────────────────────────────────────────────
# 3. AGING ANALYSIS
# ─────────────────────────────────────────────
def aging_bucket(row):
    """
    Bucket open invoices by days overdue relative to AS_OF_DATE.
    Using today's live date on historical data would push everything into
    '90+ Days Overdue', hiding the real distribution.
    """
    if pd.isna(row["due_date"]):
        return "Unknown – No Due Date"
    days = (AS_OF_DATE - row["due_date"]).days
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


open_df["aging_bucket"]  = open_df.apply(aging_bucket, axis=1)
open_df["days_overdue"]  = (AS_OF_DATE - open_df["due_date"]).dt.days.clip(lower=0)

# ─────────────────────────────────────────────
# 4. SUMMARY STATISTICS
# ─────────────────────────────────────────────
total_invoiced  = df_financial["amount"].sum()
total_collected = paid_df["amount"].sum()
total_open      = open_df["amount"].sum()
collection_rate = (total_collected / total_invoiced * 100) if total_invoiced > 0 else 0

# FIX 2 — Average days late calculated only over invoices that were actually late.
late_invoices    = paid_df[paid_df["days_late"] > 0]
avg_days_late    = late_invoices["days_late"].mean() if len(late_invoices) > 0 else 0
avg_days_early   = paid_df[paid_df["days_early"] > 0]["days_early"].mean()
pct_on_time      = (paid_df["application_status"].str.startswith("Paid – On Time").sum()
                    / len(paid_df) * 100) if len(paid_df) > 0 else 0
data_quality_cnt = paid_df["application_status"].str.startswith("Data Quality").sum()

print(f"\nFinancial Summary")
print(f"  Total Invoiced       : ${total_invoiced:>15,.2f}")
print(f"  Total Collected      : ${total_collected:>15,.2f}")
print(f"  Total Outstanding    : ${total_open:>15,.2f}")
print(f"  Collection Rate      : {collection_rate:>14.1f}%")
print(f"  % Paid On Time       : {pct_on_time:>14.1f}%")
print(f"  Avg Days Late        : {avg_days_late:>14.1f} days  (late invoices only)")
print(f"  Avg Days Early       : {avg_days_early:>14.1f} days  (early payments)")
print(f"  Data Quality Issues  : {data_quality_cnt:>14,} invoices")

print(f"\nCash Application Status Breakdown:")
print(paid_df["application_status"].value_counts().to_string())

print(f"\nAR Aging Breakdown (Open Invoices) — as of {AS_OF_DATE.strftime('%d-%b-%Y')}:")
aging_summary = open_df.groupby("aging_bucket")["amount"].agg(["count", "sum"])
aging_summary.columns = ["Count", "Amount ($)"]
aging_summary["Amount ($)"] = aging_summary["Amount ($)"].map("${:,.2f}".format)
print(aging_summary.to_string())

# ─────────────────────────────────────────────
# 5. TOP CUSTOMERS
# ─────────────────────────────────────────────
# FIX 5 — Exclude "Unknown Customer" so real accounts drive the priority list.
top_customers = (
    open_df[open_df["name_customer"] != "Unknown Customer"]
    .groupby("name_customer")["amount"]
    .sum()
    .sort_values(ascending=False)
    .head(10)
    .reset_index()
)
top_customers.columns = ["Customer Name", "Outstanding Amount ($)"]

print(f"\nTop 10 Customers by Outstanding Amount (identified customers only):")
print(top_customers.to_string(index=False))

# Count how much is under unknown customers for transparency
unknown_outstanding = open_df[open_df["name_customer"] == "Unknown Customer"]["amount"].sum()
if unknown_outstanding > 0:
    print(f"\n  ⚠ Unknown Customer total outstanding: ${unknown_outstanding:,.2f}  "
          f"(excluded from Top 10 — data quality review recommended)")

# ─────────────────────────────────────────────
# 6. PREPARE EXPORT DATA
# ─────────────────────────────────────────────
cash_app_export = paid_df[[
    "invoice_id", "name_customer", "business_code", "invoice_currency",
    "amount", "create_date", "due_date", "clear_date",
    "days_to_pay", "days_early", "days_late", "application_status"
]].copy()
cash_app_export.columns = [
    "Invoice ID", "Customer Name", "Business Code", "Currency",
    "Invoice Amount ($)", "Invoice Date", "Due Date", "Payment Date",
    "Days to Pay (signed)", "Days Paid Early", "Days Paid Late", "Application Status"
]
for col in ["Invoice Date", "Due Date", "Payment Date"]:
    cash_app_export[col] = pd.to_datetime(cash_app_export[col]).dt.strftime("%d-%b-%Y")

aging_export = open_df[[
    "invoice_id", "name_customer", "business_code", "invoice_currency",
    "amount", "due_date", "days_overdue", "aging_bucket"
]].copy()
aging_export.columns = [
    "Invoice ID", "Customer Name", "Business Code", "Currency",
    "Outstanding Amount ($)", "Due Date", "Days Overdue", "Aging Bucket"
]
aging_export["Due Date"] = pd.to_datetime(aging_export["Due Date"]).dt.strftime("%d-%b-%Y")

aging_pivot = open_df.groupby("aging_bucket")["amount"].agg(
    Count="count", Total_Amount="sum"
).reset_index()
aging_pivot.columns = ["Aging Bucket", "Invoice Count", "Total Outstanding ($)"]

top_cust_export = top_customers.copy()

data_quality_export = paid_df[
    paid_df["application_status"].str.startswith("Data Quality")
][[
    "invoice_id", "name_customer", "amount", "due_date", "clear_date", "application_status"
]].copy()
data_quality_export.columns = [
    "Invoice ID", "Customer Name", "Amount ($)", "Due Date", "Clear Date", "Issue"
]

summary_df = pd.DataFrame({
    "Metric": [
        "As-Of Reference Date",
        "Total Invoices (all rows)",
        "Total Financial Invoices (amount ≠ 0)",
        "Paid Invoices",
        "Open Invoices",
        "Total Invoiced ($)",
        "Total Collected ($)",
        "Total Outstanding ($)",
        "Collection Rate (%)",
        "% Paid On Time",
        "Avg Days Late (late invoices only)",
        "Avg Days Early (early payments)",
        "Data Quality Issues (paid set)",
        "Unknown Customer Outstanding ($)",
    ],
    "Value": [
        AS_OF_DATE.strftime("%d-%b-%Y"),
        f"{total_all:,}",
        f"{len(df_financial):,}",
        f"{len(paid_df):,}",
        f"{len(open_df):,}",
        f"${total_invoiced:,.2f}",
        f"${total_collected:,.2f}",
        f"${total_open:,.2f}",
        f"{collection_rate:.1f}%",
        f"{pct_on_time:.1f}%",
        f"{avg_days_late:.1f} days",
        f"{avg_days_early:.1f} days",
        f"{data_quality_cnt:,}",
        f"${unknown_outstanding:,.2f}",
    ]
})

# ─────────────────────────────────────────────
# 7. EXPORT TO EXCEL
# ─────────────────────────────────────────────
output_file = "O2C_Full_Report.xlsx"

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    summary_df.to_excel(writer,          sheet_name="Executive Summary",   index=False)
    cash_app_export.to_excel(writer,     sheet_name="Cash Application",    index=False)
    aging_export.to_excel(writer,        sheet_name="AR Aging Detail",     index=False)
    aging_pivot.to_excel(writer,         sheet_name="Aging Buckets",       index=False)
    top_cust_export.to_excel(writer,     sheet_name="Top 10 Outstanding",  index=False)
    data_quality_export.to_excel(writer, sheet_name="Data Quality Issues", index=False)

# ─────────────────────────────────────────────
# 8. FORMAT EXCEL
# ─────────────────────────────────────────────
STATUS_COLORS = {
    # Cash application statuses
    "Paid – On Time"                        : "C6EFCE",
    "Paid – Slightly Late (≤15 days)"       : "FFEB9C",
    "Paid – Late (16–30 days)"              : "FFCC99",
    "Paid – Very Late (>30 days)"           : "FFC7CE",
    "Data Quality Issue – No Clear Date"    : "DDEBF7",
    "Data Quality Issue – No Due Date"      : "BDD7EE",
    # Aging buckets
    "Current (Not Yet Due)"                 : "C6EFCE",
    "1–30 Days Overdue"                     : "FFEB9C",
    "31–60 Days Overdue"                    : "FFCC99",
    "61–90 Days Overdue"                    : "FFC7CE",
    "90+ Days Overdue"                      : "FF4444",
    "Unknown – No Due Date"                 : "DDEBF7",
}

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
CENTER      = Alignment(horizontal="center", vertical="center", wrap_text=True)
thin        = Side(style="thin", color="CCCCCC")
BORDER      = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = openpyxl.load_workbook(output_file)

for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]

    # Style header row
    for cell in ws[1]:
        cell.fill      = HEADER_FILL
        cell.font      = HEADER_FONT
        cell.alignment = CENTER
        cell.border    = BORDER

    # Find which column holds status/bucket values for colour-coding
    color_col = None
    for idx, cell in enumerate(ws[1], 1):
        if cell.value in ("Application Status", "Aging Bucket", "Issue"):
            color_col = idx

    # Style data rows
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

    # Auto-fit column widths
    for col in ws.columns:
        max_len = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 35)

    ws.freeze_panes = "A2"

wb.save(output_file)

print(f"\n✅ Full O2C Report saved : {output_file}")
print(f"   6 sheets generated    : Executive Summary, Cash Application,")
print(f"                           AR Aging Detail, Aging Buckets,")
print(f"                           Top 10 Outstanding, Data Quality Issues")
print("=" * 60)
