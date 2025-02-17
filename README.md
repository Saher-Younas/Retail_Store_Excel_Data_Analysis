# Retail_Store_Sales_Excel_Data_Analysis

## Overview
This repository contains an Excel-based **Sales Data Analysis & Dashboard** designed to track key business metrics, generate insights, and provide interactive visualizations using pivot tables, slicers, and charts. The Excel file combines **data cleaning, statistical analysis, and dynamic dashboards** in a single workbook.

---

## How to Download & Extract Files
1. Click on the ZIP file link: [Retail_Sales_Full_Project_Dashboard.zip](https://github.com/Saher-Younas/Retail_Store_Excel_Data_Analysis/blob/main/Retail_Sales_Full_Project_Dashboard.zip).
2. Download the ZIP file to your computer.
3. Extract the contents using one of the following methods:
   - **Windows**: Right-click the file and select "Extract All."
   - **Mac**: Double-click the ZIP file to unzip it.
   - **Linux**: Use the command `unzip Sales_Analysis_Resources.zip` in the terminal.
4. Open the extracted Excel files in Microsoft Excel to explore the data and dashboard.
5. Ensure that macros and external links are enabled for full functionality.
6. When prompted, **enable macros** so that the automated features (data entry form, pivot table refreshes) function correctly.

---

## Data Cleaning and Preprocessing
### **Raw Data (Sheet: "Raw Data")**
- Contains original sales transactions (dates, categories, payment methods, costs, revenues).
- May have duplicates or blank cells before cleaning.

### **Cleaned Data (Sheet: "Cleaned Data")**
- **Duplicates removed** to preserve data accuracy.
- **Blank cells handled** by filling with the **mean** of each column.
- **Formats standardized** (date, numeric types) for consistent analysis.
- **Created New Columns for Enhanced Analysis:**
  - **Year, Month, Day, Date**: Extracted from the **Order Date** column using the `TEXT()` formula.
  - **Delivery Time**: Calculated using the difference between Order and Delivery dates.
  - **Total Cost**: Derived from the relevant cost calculations (specific formula used).
  - **Sales Revenue**: Obtained by multiplying **Quantity** with **Unit Price**.
  - **Net Profit**: Computed as **Sales Revenue - Total Cost**.

This ensures all subsequent analyses and visualizations are based on **reliable, uniform data**.

---

## Statistical Analysis
### **Data Analysis (Sheet: "Data_Analysis")**
- Contains **descriptive statistics** (mean, median, standard deviation, range) for key variables: **Delivery Time, Total Cost, Sales Revenue, Net Profit, Quantity, and Unit Price**.
- Includes a **T-Test** to examine the relationship between **delivery time and return status** (Completed vs. Returned).
- **Results:**
  - **p-value < 0.05**, indicating **longer delivery times correlate with higher return likelihood**.
  
### **Business Recommendations:**
1. Closely **monitor orders exceeding a 7-day delivery window**.
2. Implement **alerts for delayed deliveries** to reduce return rates.

---

## KPI (Sheet: "KPI")
- Summarizes **key performance indicators (KPIs)** using pivot tables derived from the **Cleaned Data**.
- Displays metrics such as **Total Revenue, Total Cost, Net Profit, and Product-Level Summaries**.
- Provides a concise **overview of performance** for quick decision-making.

---

## Sales Form (Sheet: "Sales Form")
- A **user-friendly data entry form** for adding new sales records.
- Uses **data validation** to ensure accuracy (dropdown lists for categories, payment methods).
- Employs **macros** to **automate data transfer and pivot table refreshes**.

By updating the workbook **in real time**, any new sales data is immediately **reflected in reports and dashboards**.

---

## Interactive Dashboard (Sheet: "Dashboard")
- Offers **real-time insights into sales performance**, featuring:
  - **Revenue by Country** (dynamic map)
  - **Sales by Category** (bar charts)
  - **Orders by Payment Method** (pie chart)
  - **Daily Revenue Trends** (line charts)
  - **Slicers** for filtering data by country and category

### **Navigation Icons**
On the left side of the dashboard, six clickable icons link to:
1. **Refresh Data** – Updates all pivot tables and calculations.
2. **Dashboard** – Returns to the main interface.
3. **Descriptive Statistics** – Opens the Data_Analysis sheet for summary metrics and T-tests.
4. **Sales Form** – Jumps to the data entry form.
5. **Cleaned Dataset** – Displays the processed, duplicate-free data.
6. **Map Visualization** – Highlights **sales distribution** across regions.

---

## Map Visualization (Sheet: "Map_Visualization")
- Provides a **geographic breakdown of sales revenue** by country.
- Linked to **dashboard slicers**, allowing **region-specific analysis**.
- **Resolved slicer-connection issues** by ensuring Report Connection and Data Model consistency.

---

## File Structure
- **SalesDashboard.xlsm**
  - **Raw Data** – Original dataset with potential duplicates or missing values
  - **Cleaned Data** – Refined table for accurate calculations
  - **Data Analysis** – Statistical summaries (descriptive statistics, T-test)
  - **KPI** – Key performance metrics derived from pivot tables
  - **Sales Form** – Macro-enabled data entry interface
  - **Dashboard** – Main visual reporting sheet with slicers
  - **Map Visualization** – Geographic view of sales distribution

---

## Contributing
If you have **improvements or suggestions**, feel free to **open an issue or submit a pull request**. This workbook can be adapted to **various datasets or industries** by modifying the raw data and updating pivot table connections.

---

## Conclusion
By combining **data cleaning, statistical analysis, pivot tables, and an interactive dashboard** in **SalesDashboard.xlsm**, this project demonstrates how **Excel can serve as a powerful platform for business intelligence**. From **quick data entry** to **in-depth performance metrics**, the entire workflow is consolidated into a **single, user-friendly file** that supports **real-time decision-making**.

