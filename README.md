# Super Store Orders - Advanced Excel Data Analysis & Interactive Dashboard

A comprehensive data analysis and visualization project built using **Microsoft Excel**. This project explores a global superstore dataset to analyze sales performance, profitability, customer segmentation, and shipping efficiency to extract actionable business insights.

---

## 📌 Project Overview

A global superstore sells **Furniture, Office Supplies and Technology** products to three customer segments across **7 markets and 147 countries**. Revenue is growing, but not every sale makes money. This project digs into the numbers to find out:

- **What** is driving profit, and what is quietly losing money?
- **Where** (market and country) is the business strong or weak?
- **Who** (which customer segment) brings the most value?
- **How** do discounts and shipping choices affect results?

Everything is done inside a single Excel workbook, so anyone can open it, refresh the PivotTables and explore the results.

---

| Item | Detail |
|---|---|
| **Source** | [ https://www.kaggle.com/datasets/laibaanwer/superstore-sales-dataset) – 
| **File used** | `https://1drv.ms/x/c/7e10cac2ae4d4283/IQDST9-Nzk8FSI_UexpCSjFNAWtB9yvUxx71A-BTfmeh9uE?e=eG6WuY` (CSV converted to Excel) |
| **Rows × Columns** | 51,290 rows × 22 columns (21 original + 1 calculated) |
| **Time period** | 1 Jan 2011 – 31 Dec 2014 |
| **Coverage** | 7 markets, 13 regions, 147 countries, 795 customers, 10,292 products |
| **Granularity** | One row = one product line inside an order (25,035 unique orders) |

> 🙏 Data credit goes to the dataset author on Kaggle. This repository is for learning and portfolio purposes only.

## 🧰 Tools & Skills Used

- **Microsoft Excel** – Excel Tables (structured references), PivotTables, PivotCharts, Slicers/Filters
- **Formulas** – `IFERROR`, calculated columns, averages and percentage-difference ("Show Values As")
- **Data cleaning** – null/duplicate/consistency checks, data-type validation
- **Dashboard design** – combining multiple PivotCharts into one view
- **Business analysis** – turning numbers into insights and recommendations

---

### Data Dictionary

## 🧹 Data Cleaning & Preparation

| Step | What was done | Result |
|---|---|---|
| 1. Import | Loaded the Kaggle CSV into Excel and saved as `.xlsx` | 51,290 rows, 22 columns |
| 2. Convert to Table | Turned the range into an Excel Table (`Table1`) so formulas and PivotTables update automatically | ✅ |
| 3. Header check | Column names kept in consistent `snake_case` | ✅ |
| 4. Missing values | Checked every column | **0 missing values** |
| 5. Duplicates | Checked full-row duplicates | **0 duplicate rows** |
| 6. Date validation | Confirmed dates are stored as real Excel dates and `ship_date ≥ order_date` | **0 invalid records** |
| 7. Value checks | Checked sales, quantity, discount ranges | Quantity ≥ 1, discount between 0 and 0.85 |
| 8. Feature engineering | Added `Profit margin = IFERROR(profit / sales, 0)` | Avoids divide-by-zero errors |
| 9. Year field | Used a `year` column for time-based analysis | Enables YoY growth |

> 📝 *Note:* `order_id` repeats because one order can contain multiple products. That is expected, so these rows were **not** treated as duplicates.

---
## ❓ Business Questions & Analysis Process

### Q1 (i) – Which categories and sub-categories generate the most sales and profit?
- **Process:** PivotTable with `category → sub_category` in rows; values = Sum of `sales`, Sum of `profit`, Average of `Profit margin`.
- **Result:** Technology earns the most profit (**$663.8K**), followed by Office Supplies (**$518.5K**) and Furniture (**$286.8K**). *Copiers* is the top sub-category by profit (**$258.6K**), and *Paper* has the best average margin (**19.7%**).

### Q1 (ii) – Which sub-category is losing money?
- **Process:** Filtered the pivot to the sub-category with negative profit.
- **Result:** **Tables** generated **$757K in sales** but a **$64.1K loss** (average margin **−24.2%**), the only sub-category with a negative total profit.

### Q1 (iii) – How do discounts relate to profit margin?
- **Process:** PivotTable with Sum of `discount`, Average of `Profit margin` and Sum of `profit`.
- **Result:** Overall average margin is **4.7%** across all order lines. Discounts are studied further in the findings below.

### Q2 – How has sales grown year over year?
- **Process:** PivotTable by `year`, Sum of `sales`, *Show Values As → % Difference From (previous year)*.
- **Result:** Growth of **+18.5% (2012)**, **+27.2% (2013)** and **+26.3% (2014)**. Sales nearly doubled from **$2.26M (2011)** to **$4.30M (2014)**.

### Q3 (i) – Which customer segment is most valuable?
- **Process:** PivotTable by `segment` with Sum of sales, Sum of profit, Average margin and Count of order lines.
- **Result:**

| Segment | Sales | Profit | Order lines |
|---|---|---|---|
| Consumer | $6.51M | $749.2K | 26,518 |
| Corporate | $3.82M | $442.8K | 15,429 |
| Home Office | $2.31M | $277.0K | 9,343 |

### Q3 (ii) – Do segments differ in average order-line value?
- **Process:** PivotTable with Average of `sales` by `segment`.
- **Result:** Almost identical: **Consumer $245.4**, **Corporate $247.9**, **Home Office $247.2**. Segments differ in *volume*, not in *spend per line*.

### Q4 (i) – Are markets/countries profitable after shipping cost?
- **Process:** PivotTable with `market → country`, Sum of `profit` and Sum of `shipping_cost`.
- **Result (Africa):** Profit is **$88.9K** but shipping cost is **$88.1K**, so shipping nearly equals profit. **Nigeria** loses **$80.8K** and drags the whole market down.

### Q4 (ii) – Which shipping mode is used most and how fast is it?
- **Process:** PivotTable by `ship_mode` with Sum of `sales` and a delivery-time comparison.
- **Result:** **Standard Class** carries **~60%** of sales, followed by Second Class (~20%), First Class (~14%) and Same Day (~5%).

---

## 📈 Excel Dashboard & Features
An **Executive Interactive Dashboard** tab was created to consolidate macro metrics and visual insights:
* **Key Performance Indicators (KPIs):** High-level summary metrics showing Total Sales, Total Profit, Total Units Sold, and Average Profit Margin.
* **Data Visualizations:** Regional revenue distribution charts, category breakdowns, and yearly trendlines.
* **Interactive Design:** Integrated Pivot Charts and conditional formatting to visually highlight key financial thresholds.

---

## 🔎 Key Findings

1. **Overall performance:** Total sales are **$12.64M** and total profit is **$1.47M**, a weighted profit margin of **about 11.6%**.
2. **Strong growth:** Sales rose in every year, up **~90% from 2011 to 2014**. Profit grew from **$249K to $504K** over the same period.
3. **Technology leads profit:** Technology brings **~45%** of total profit, and Copiers and Phones are its biggest contributors.
4. **Tables are a problem:** Tables is the only sub-category with a net loss (**−$64.1K**) despite $757K in sales. Machines, Appliances and Binders also have slightly negative average margins.
5. **Paper and Labels are the most efficient:** They have the best average margins (**~19.7%** and **~11.9%**), though with smaller sales volume.
6. **Segments behave alike:** Consumer is the largest segment (~51% of sales), but the three segments have similar margins and spend per order line.
7. **Market gaps:** APAC is the biggest market by profit (**$437.6K**), followed by EU (**$372.8K**) and US (**$286.4K**). Canada is small but healthy; Africa and EMEA earn little relative to their sales.
8. **Country red flags:** Turkey (**−$98.4K**), Nigeria (**−$80.8K**), Netherlands (**−$41.1K**), Honduras (**−$29.5K**) and Pakistan (**−$22.4K**) are the weakest countries by profit.
9. **Discounts hurt profit:** Order lines with **no discount** produced about **$1.77M profit**, while lines discounted **above 40%** produced a loss of about **$627K** and lines discounted **20–40%** lost about **$187K**. *(Supplementary check; see Limitations.)*
10. **Shipping speed:** Standard Class takes about **5 days**, Second Class **3.2 days**, First Class **2.2 days** and Same Day **under 1 day** on average.

---


## ⚠️ Project Limitations
* **Static Historical Data:** The analysis is based on a static historical snapshot and does not capture real-time market changes or dynamic pricing.
* **Lack of Behavioral Context:** The dataset contains transactional metrics but lacks granular customer behavioral attributes (e.g., website traffic, marketing campaign conversion rates).
* **External Factors:** External macroeconomic variables—such as inflation rates, currency exchange fluctuations, and local tariffs—are not represented.

---

## 🚀 Recommended Next Steps & Future Enhancements
1. **RFM Analysis:** Implement Recency, Frequency, and Monetary (RFM) customer segmentation to identify high-value customer clusters.
2. **BI Tool Migration:** Replicate and expand the data model in Power BI or Tableau for automated data refreshing and enhanced interactive filtering.
3. **Python Integration:** Perform exploratory data analysis (EDA) using Python (Pandas, Seaborn, Matplotlib) to perform predictive sales forecasting.

---

## 👨‍💻 Author
**Nusrat Jahan Fariha**  
*Department of Statistics and Data Science*  
*Jahangirnagar University*  

* **LinkedIn:** https://www.linkedin.com/in/nusrat-jahan-fariha-83a60a445/?isSelfProfile=true
* **GitHub:** https://github.com/nusratjahanfariha110-dotcom
---
*If you find this project useful, feel free to give it a ⭐ star on GitHub!***
