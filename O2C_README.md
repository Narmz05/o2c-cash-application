# O2C Cash Application & Invoice Reconciliation

I built this project to understand how Order to Cash works in real finance operations teams. The dataset has 50,000 real invoice records and I used Python to automate the cash application and reconciliation process that finance teams do manually every day.

---

## Why I Built This

While learning about O2C processes, I noticed that matching payments to invoices and tracking overdue accounts is something finance teams spend a lot of time doing manually. So I decided to automate it using Python and see what insights I could pull from a real dataset.

---

## What the Project Does

- Loads 50,000 invoice records and separates paid vs open invoices
- Applies payments to invoices and classifies each as paid on time, late, or very late
- Groups overdue invoices into aging buckets (1–30, 31–60, 61–90, 90+ days)
- Finds the top 10 customers with the highest outstanding amounts
- Exports everything into a formatted Excel report with 5 sheets

---

## What I Found (Results)

| Metric | Value |
|--------|-------|
| Total Invoices | 50,000 |
| Total Invoiced | $1,616,851,082 |
| Total Collected | $1,282,601,395 |
| Total Outstanding | $334,249,686 |
| Collection Rate | 79.3% |

Most of the overdue amount (around $159M) was sitting in the 61–90 days bucket which tells me those accounts need urgent collections follow-up.

---

## Files in This Repo

```
o2c-cash-application/
├── dataset.csv                 # 50,000 invoice records
├── cash_application_real.py    # Main Python script
├── O2C_Full_Report.xlsx        # Output report (5 sheets)
└── README.md
```

---

## Excel Report Sheets

- **Executive Summary** — overall KPIs at a glance
- **Cash Application** — all 40,000 paid invoices with payment status
- **AR Aging Detail** — 10,000 open invoices with aging classification
- **Aging Buckets** — summary table by aging bucket
- **Top 10 Outstanding** — customers to prioritize for collections

---

## How to Run

```bash
# Install dependencies
pip install pandas openpyxl numpy

# Run the script
python cash_application_real.py
```

The Excel report will be generated automatically in the same folder.

---

## Tools Used

- Python, Pandas, NumPy, OpenPyXL

---

## About Me

I'm Narmadhadevi, a final year B.Tech IT student at Anna University. I built this as part of learning Order to Cash finance operations concepts.

📧 narmadhadevi1008@gmail.com  
🔗 linkedin.com/in/narmadhadevi