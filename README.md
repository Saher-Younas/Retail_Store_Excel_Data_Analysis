# Retail_Store_Sales_Excel_Data_Analysis

## Overview
This repository contains an Excel-based **Sales Data Analysis & Dashboard** designed to track key business metrics, generate insights, and provide interactive visualizations using pivot tables, slicers, and charts. The Excel file combines **data cleaning, statistical analysis, and dynamic dashboards** in a single workbook.

---

## How to Download & Extract Files
1. Download the required files
Raw Data : [Sales_Raw_Data.xlsx](https://github.com/Saher-Younas/Retail_Store_Excel_Data_Analysis/blob/main/sales_raw_data.xlsx)                                                       Cleaned Data : [Retail_Store_Sales_Cleaned_Data.xlsx](https://github.com/Saher-Younas/Retail_Store_Excel_Data_Analysis/blob/main/Retail_Store_Sales_Cleaned_Data.xlsx)
   - Locate the downloaded XLSX file on your computer.
   - Double-click to open it in Microsoft Excel (or an alternative spreadsheet software like Google Sheets).
2. Click on the ZIP file link: [Retail_Sales_Full_Project_Dashboard.zip](https://github.com/Saher-Younas/Retail_Store_Excel_Data_Analysis/blob/main/Retail_Sales_Full_Project_Dashboard.zip).
3. Download the ZIP file to your computer.
4. Extract the contents using one of the following methods:
   - **Windows**: Right-click the file and select "Extract All."
   - **Mac**: Double-click the ZIP file to unzip it.
   - **Linux**: Use the command `unzip Sales_Analysis_Resources.zip` in the terminal.
5. Open the extracted Excel files in Microsoft Excel to explore the data and dashboard.
6. Ensure that macros and external links are enabled for full functionality.
7. When prompted, **enable macros** so that the automated features (data entry form, pivot table refreshes) function correctly.

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

  ![Data_Analysis](https://github.com/user-attachments/assets/ff8e3fd9-fd62-446f-9c68-be44dc5c10ac)

  
### **Business Recommendations:**
1. Closely **monitor orders exceeding a 7-day delivery window**.
2. Implement **alerts for delayed deliveries** to reduce return rates.

---

## KPI (Sheet: "KPI")
- Summarizes **key performance indicators (KPIs)** using pivot tables derived from the **Cleaned Data**.
- Displays metrics such as **Total Revenue, Total Cost, Net Profit, and Product-Level Summaries**.
- Provides a concise **overview of performance** for quick decision-making.

  ![image](https://github.com/user-attachments/assets/ed788da8-6a54-43ce-9314-237049569202)


---

## Sales Form (Sheet: "Sales Form")

- A **user-friendly data entry form** for adding new sales records.
- Uses **data validation** to ensure accuracy (dropdown lists for categories, payment methods).
- Employs **macros** to **automate data transfer and pivot table refreshes**.
  
![Sales_Form](https://github.com/user-attachments/assets/65d31744-21f3-46b3-8837-27db5a1253b5)

By updating the workbook **in real time**, any new sales data is immediately **reflected in reports and dashboards**.

---

## Interactive Dashboard (Sheet: "Dashboard")
- Offers **real-time insights into sales performance**, featuring:

![Dashboard](https://github.com/user-attachments/assets/534d1b7f-a441-4acd-9af7-4a1074707f9d)


  - **Revenue by Country** (dynamic map)
  - **Sales by Category** (bar charts)
  - **Orders by Payment Method** (pie chart)
  - **Daily Revenue Trends** (line charts)
  - **Slicers** for filtering data by country and category

### **Navigation Icons**
On the left side of the dashboard, six clickable icons link to:



1. **Refresh Data** ![image](https://github.com/user-attachments/assets/fa95ccb6-fbf0-48b3-8c66-3ef3ca3d4323)– Updates all pivot tables and calculations.
2. **Dashboard** ![image](https://github.com/user-attachments/assets/02b17cef-56af-44fa-b7ec-7537e085d7bd)
  – Returns to the main interface.
3. **Descriptive Statistics** ![image](https://github.com/user-attachments/assets/c6df1649-1cb7-432d-b57a-e40107b39bea)
  – Opens the Data_Analysis sheet for summary metrics and T-tests.
4. **Sales Form**  ![image](https://github.com/user-attachments/assets/cab73641-5f2a-41ec-aed5-c1ee9770c2b0)
 – Jumps to the data entry form.
5. **Cleaned Dataset** ![image](https://github.com/user-attachments/assets/3d9f48cd-5f98-4bab-84ca-af87e6a36964)
 – Displays the processed, duplicate-free data.
6. **Map Visualization** ![image](https://github.com/user-attachments/assets/33d05d92-cd96-471d-bc28-50193c02a12c)
 – Highlights **sales distribution** across regions.

---

## Map Visualization (Sheet: "Map_Visualization")
- Provides a **geographic breakdown of sales revenue** by country.
- Linked to **dashboard slicers**, allowing **region-specific analysis**.
- **Resolved slicer-connection issues** by ensuring Report Connection and Data Model consistency.

---

## File Structure
- **Retail_Sales_Full_Project_Dashboard.xlsm**
  - **Cleaned Data** – Refined table for accurate calculations and named as Retail Store Sale
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
By combining **data cleaning, statistical analysis, pivot tables, and an interactive dashboard** in **Retail_Sales_Full_Project_Dashboard.xlsm**, this project demonstrates how **Excel can serve as a powerful platform for business intelligence**. From **quick data entry** to **in-depth performance metrics**, the entire workflow is consolidated into a **single, user-friendly file** that supports **real-time decision-making**.

