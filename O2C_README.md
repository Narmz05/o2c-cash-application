# 💰 Order to Cash – Cash Application & Invoice Reconciliation System

A Python-based end-to-end **Order to Cash (O2C)** automation project built on a real-world dataset of **50,000 invoice records**. Automates cash application, invoice reconciliation, AR aging analysis, and generates a complete multi-sheet Excel report — replicating real O2C finance operations workflows.

---

## 📊 Dataset

- **Source:** Real-world O2C Invoice Dataset (Kaggle)
- **Records:** 50,000 invoices
- **Scope:** Multi-business, multi-currency (USD & CAD)
- **Fields:** Invoice ID, Customer, Business Code, Amount, Due Date, Clear Date, Payment Terms, Open/Closed Status

---

## 🎯 What This Project Does

| Module | Description |
|--------|-------------|
| **Cash Application** | Matches cleared payments to invoices, classifies as On Time / Late / Very Late |
| **AR Reconciliation** | Separates open vs closed invoices, calculates outstanding balances |
| **Aging Analysis** | Buckets open invoices into Current / 1–30 / 31–60 / 61–90 / 90+ days overdue |
| **Collections Report** | Identifies top 10 customers by outstanding amount |
| **Executive Summary** | KPI dashboard — collection rate, total invoiced, collected, outstanding |

---

## 📋 Key Results (from 50,000 records)

| Metric | Value |
|--------|-------|
| Total Invoices | 50,000 |
| Total Invoiced | $1,616,851,082 |
| Total Collected | $1,282,601,395 |
| Total Outstanding | $334,249,686 |
| Collection Rate | **79.3%** |
| Paid On Time | 23,236 invoices |

### AR Aging Breakdown
| Bucket | Invoices | Amount |
|--------|----------|--------|
| Current | 8 | $235,790 |
| 1–30 Days Overdue | 104 | $1,399,002 |
| 31–60 Days Overdue | 2,906 | $92,470,574 |
| 61–90 Days Overdue | 4,549 | $159,104,141 |
| 90+ Days Overdue | 2,433 | $81,040,176 |

---

## 📁 Project Structure

```
o2c-cash-application/
│
├── dataset.csv                  # Real-world O2C invoice dataset (50,000 records)
├── cash_application.py          # Main automation script
├── O2C_Full_Report.xlsx         # Auto-generated 5-sheet Excel report
└── README.md
```

---

## 📋 Excel Report Sheets

| Sheet | Contents |
|-------|---------|
| **Executive Summary** | KPI dashboard — collection rate, totals |
| **Cash Application** | 40,000 paid invoices with application status |
| **AR Aging Detail** | 10,000 open invoices with aging classification |
| **Aging Buckets** | Pivot summary by aging bucket |
| **Top 10 Outstanding** | Highest outstanding customers for collections follow-up |

---

## 🛠️ Tools & Technologies

| Tool | Purpose |
|------|---------|
| Python 3.x | Core automation |
| Pandas | Data processing & reconciliation |
| NumPy | Numerical operations |
| OpenPyXL | Excel report generation & formatting |

---

## 🚀 How to Run

```bash
# 1. Clone the repository
git clone https://github.com/yourusername/o2c-cash-application.git
cd o2c-cash-application

# 2. Install dependencies
pip install pandas openpyxl numpy

# 3. Run the script
python cash_application.py

# 4. Open the report
O2C_Full_Report.xlsx
```

---

## 💡 Key O2C Concepts Demonstrated

- **Cash Application** — Matching incoming payments to open invoices
- **AR Aging** — Classifying overdue receivables by time buckets
- **Collections Management** — Identifying priority accounts for follow-up
- **Reconciliation** — Balancing total invoiced vs collected amounts
- **Financial Reporting** — Generating ERP-style O2C dashboards

---

## 👩‍💻 Author

**Narmadhadevi C**
B.Tech Information Technology | Anna University
📧 narmadhadevi1008@gmail.com
🔗 LinkedIn/narmadhadevi
