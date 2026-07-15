### Project Video
https://github.com/user-attachments/assets/fb896dcb-3229-4b37-b964-1e99c3cccc39

### Advanced Microsoft Excel Dashboard Masterclass – Adventure Works Sales Analysis

A production-ready Microsoft Excel Business Intelligence (BI) dashboard built using **Power Query**, **Power Pivot**, **Data Modeling**, and **DAX (Data Analysis Expressions)**.

This project demonstrates how to transform a relational database into a modern, interactive sales dashboard without relying on traditional Excel lookup formulas. It follows industry best practices for data preparation, modeling, analytics, and dashboard design.

---

## Project Overview

The goal of this project is to build a professional **Adventure Works Sales Dashboard** capable of analyzing sales performance through an interactive and scalable reporting system.

Instead of using complex worksheet formulas such as **VLOOKUP**, **XLOOKUP**, or **INDEX/MATCH**, the project leverages Excel's Business Intelligence tools to create a faster, cleaner, and more maintainable solution.

## Key Objectives

* Build a professional Business Intelligence dashboard in Microsoft Excel
* Create a scalable data model using Power Pivot
* Clean and transform raw data with Power Query
* Perform advanced calculations using DAX
* Compare Year-over-Year (YoY) business performance
* Design a modern and user-friendly dashboard interface

---

## Features

* Interactive dashboard with slicers
* Dynamic KPI cards
* Year-over-Year (YoY) comparison
* Revenue, Profit, Cost, and Sales analysis
* Product and Customer performance insights
* Time-based filtering (Year & Month)
* Country-wise sales analysis
* Responsive charts and visualizations
* Professional dashboard layout
* Optimized Data Model for better performance

---

## Technologies Used

| Tool            | Purpose                                    |
| --------------- | ------------------------------------------ |
| Microsoft Excel | Dashboard Development                      |
| Power Query     | Data Extraction, Cleaning & Transformation |
| Power Pivot     | Data Modeling                              |
| DAX             | Business Calculations & Measures           |
| Pivot Tables    | Data Aggregation                           |
| Pivot Charts    | Data Visualization                         |
| Slicers         | Interactive Filtering                      |

---

## Project Structure

```text
Adventure-Works-Dashboard/
│
├── Dashboard.xlsx
├── README.md
│
├── Data/
│   ├── Adventure Works.xlsx
│   ├── FactInternetSales
│   ├── DimDate
│   ├── DimProduct
│   ├── DimCustomer
│   └── Other Dimension Tables
│
└── Images/
    ├── Icons
    ├── Backgrounds
    └── UI Assets
```

---

## Data Architecture

The project follows a **Star Schema** design.

## Fact Table

* **FactInternetSales**

  * Sales Amount
  * Order Quantity
  * Product Cost
  * Order Date
  * Customer Keys

## Dimension Tables

* **DimDate**
* **DimProduct**
* **DimCustomer**
* **DimGeography**
* **Other Supporting Dimensions**

Relationships are created inside **Power Pivot**, eliminating the need for repetitive lookup formulas.

---

### Implementation Workflow

## Phase-1 Data Import & Cleaning

Import the Adventure Works dataset using **Power Query**.

Tasks performed:

* Import Excel tables
* Remove unnecessary columns
* Handle missing values
* Correct data types
* Rename columns
* Prepare clean datasets

## Recommended Loading Method

Instead of using **Close & Load**, use:

```
Close & Load To...
✔ Only Create Connection
✔ Add this data to the Data Model
```

This keeps the workbook lightweight and improves performance.

---

## Phase-2 Data Modeling

Create relationships inside **Power Pivot** using **Diagram View**.

Example relationships:

```
<img width="1466" height="779" alt="image" src="https://github.com/user-attachments/assets/51de85fb-73f7-495d-b3e5-f1538fe3854c" />
```

This Star Schema enables efficient filtering and reporting across all dimensions.

---

## Phase-3 DAX Measures

Business metrics are calculated using **DAX Measures**.

Examples include:

* Total Revenue
* Total Sales
* Total Cost
* Total Profit
* Profit Margin
* Year-over-Year Growth
* Previous Year Revenue
* Dynamic Titles
* Percentage Contribution

Example:

```excel
="Revenue from " & SELECTEDVALUE(DimDate[CalendarYear])
```

These measures automatically update based on slicer selections.

---

## Phase-4 Dashboard Design

The dashboard is built with a clean, application-style interface.

## Dashboard Components

* KPI Cards
* Interactive Slicers
* Line Charts
* Bar Charts
* Revenue Trend Analysis
* Customer Insights
* Product Performance
* Regional Sales Analysis

To improve readability:

* Hide Gridlines
* Hide Headings
* Use consistent spacing
* Apply a modern color palette
* Organize content into logical sections

---

## Dashboard KPIs

The dashboard tracks key business metrics, including:

* Total Revenue
* Total Sales
* Total Orders
* Total Profit
* Cost of Goods Sold (COGS)
* Profit Margin
* Year-over-Year Growth
* Customer Distribution
* Product Performance
* Regional Performance

---

## Dashboard Interactivity

The dashboard provides an interactive user experience through:

* Year Slicer
* Month Slicer
* Country Slicer
* Product Filtering
* Dynamic KPI updates
* Interactive Pivot Charts
* Drill-down analysis
* Dynamic titles and captions

Every visual updates automatically based on the selected filters.

---

## Error Handling

## Division by Zero

Prevent calculation errors using:

```excel
=IFERROR(Formula,0)
```

This ensures empty calculations return **0** instead of `#DIV/0!`.

---

## Multi-Selection Handling

Use context-aware DAX functions such as:

* SELECTEDVALUE()
* HASONEVALUE()
* VALUES()

These functions help maintain reliable calculations when multiple slicer selections are made.

---

## Performance Optimization

To keep the workbook fast and efficient:

* Use Data Model instead of worksheet tables
* Minimize volatile formulas
* Prefer DAX Measures over calculated columns when possible
* Load data as Connections Only
* Avoid duplicate datasets
* Use a Star Schema
* Refresh data only when necessary

---

## Business Insights

The dashboard enables users to answer important business questions such as:

* Which year generated the highest revenue?
* Which products performed best?
* Which customer segments generated the most sales?
* Which countries contributed the highest revenue?
* How has sales performance changed over time?
* What is the current profit margin?
* How does this year's performance compare with last year?

---

---

## Dashboard Preview

<img width="1103" height="653" alt="image" src="https://github.com/user-attachments/assets/e45789f6-0610-4759-91ff-0b0128335faf" />

<img width="1247" height="576" alt="image" src="https://github.com/user-attachments/assets/d229d666-4b41-4e0a-9e85-2d2d64015332" />

<img width="1054" height="185" alt="image" src="https://github.com/user-attachments/assets/a792e6af-0b10-420b-8b49-fd5295b59a97" />

---

## Learning Outcomes

* Power Query
* Power Pivot
* Data Modeling
* Star Schema Design
* DAX
* KPI Development
* Interactive Dashboards
* Business Intelligence Reporting
* Dashboard Design Best Practices

---
# Contributing
Any Suggestions Or improvements are always welcome.
