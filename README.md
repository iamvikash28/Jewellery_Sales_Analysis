# 💍 Jewellery Sales Analysis — FY 2024

A complete end-to-end Data Analyst portfolio project analyzing
jewellery sales across Gold, Silver, and Diamond categories
for the financial year 2024.

---

## 📊 Project Overview

| Detail | Info |
|--------|------|
| Domain | Retail / Jewellery |
| Period | FY 2024 (Jan–Dec) |
| Records | 500+ transactions |
| Cities | 10 Indian cities |
| Categories | Gold, Silver, Diamond |

---

## 🛠️ Tools Used

| Tool | Purpose |
|------|---------|
| Python (openpyxl) | Data generation & Excel automation |
| Microsoft Excel | Data cleaning, pivot tables, formulas |
| SQL (SQLite) | Business queries & aggregations |
| Power BI | Interactive dashboard (3 tabs) |
| GitHub | Version control & portfolio hosting |

---

## 📁 Project Structure

jewellery-sales-analysis/
│
├── jewellery_sales.py           # Python script to generate Excel data
├── Jewellery_Sales_Analysis.xlsx # Main Excel file (5 sheets)
├── jewellery_sales.csv          # Exported CSV for SQL analysis
├── jewellery_dashboard.pbix     # Power BI dashboard file
├── jewellery_dashboard.pdf      # Exported dashboard PDF
└── README.md                    # Project documentation

---

## 📈 Dashboard Pages

### 1. Overview
- Total Revenue, Orders, Avg Order Value KPIs
- Monthly Revenue Trend (line chart)
- Revenue by Category (donut chart)
- Monthly Comparison by Category (grouped bar)

### 2. Products
- Top 10 Products by Revenue (horizontal bar)
- Units Sold by Category
- Avg Order Value by Category

### 3. Customers
- Orders by Customer Type (pie chart)
- Payment Method Split (donut chart)
- City-wise Revenue (bar chart)

---

## 🔍 Key Insights

1. **Festive Season Peak** — Oct & Nov drive highest revenue
   due to Diwali demand surge
2. **Diamond dominates revenue** — contributes ~65% of total
   revenue despite fewer orders
3. **Diamond AOV is 4x higher** than Silver avg order value
4. **UPI is the most popular** payment method
5. **Mumbai & Delhi** are the top revenue-generating cities
6. **Jun–Jul is the off-season** — lowest sales months

---

## 💡 SQL Queries Covered

- Monthly revenue trend
- Revenue by category with market share %
- Top 10 best-selling products
- Monthly cross-tab (pivot in SQL)
- Customer type analysis
- City-wise revenue
- Payment method breakdown
- Quarter-over-Quarter comparison
- High-value transactions (> ₹50,000)

---

## ⚙️ How to Run

1. Install Python dependency:
   ```
   pip install openpyxl
   ```

3. Run the data generation script:
   ```
   python jewellery_sales.py
   ```
4. Open "Jewellery_Sales_Analysis.xlsx" in Excel
5. Import CSV into SQLite / DB Browser and run SQL queries
6. Open "jewellery_dashboard.pbix" in Power BI Desktop

---

## 👤 Author

**Vikash Verma**
Aspiring Data Analyst | Excel · SQL · Power BI · Python

---

*This project was built as a portfolio project to demonstrate
data analysis skills for a fresher data analyst role.*
