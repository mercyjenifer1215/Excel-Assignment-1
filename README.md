# 📊 Excel Data Exploration & Analysis

## 📌 Project Overview

This project is part of my learning journey as an aspiring Data Analyst.

The objective of this assignment is to perform basic data exploration and analysis using Microsoft Excel. The dataset contains product information such as Product ID, Product Name, Brand Name, Quantity, Category, and Price.

Through this project, I practiced essential Excel functions used in data analysis, including aggregation, logical functions, conditional calculations, and text manipulation.

This project helped me build a strong foundation in data exploration and prepare data for further analysis.

---

## 🎯 Project Objectives

The main objectives of this project are:

- Explore and understand product dataset information.
- Perform basic statistical calculations using Excel.
- Apply logical functions to categorize products.
- Use conditional aggregation functions for analysis.
- Extract useful information from Product IDs.
- Develop practical Excel skills for data analytics.

---

## 🗂️ Dataset Description

The dataset contains information about different products.

| Column | Description |
|---|---|
| Product ID | Unique identifier for each product |
| Product Name | Name of the product |
| Brand Name | Brand associated with the product |
| Quantity | Quantity of products |
| Category | Product category |
| Price ($) | Price of the product |

### Dataset Summary

- **Total Products:** 33
- **Number of Attributes:** 6 original product attributes
- **Tool Used:** Microsoft Excel
- **Dataset Type:** Product Dataset

---

## 🛠️ Tools & Technologies

- Microsoft Excel
- Excel Formulas & Functions
- Data Exploration
- Logical Functions
- Conditional Aggregation
- Text Manipulation
- GitHub

---

## 📋 Tasks Performed

### 1️⃣ Basic Data Exploration

Used Excel functions to calculate:

- Total price of all products.
- Number of products in the dataset.
- Average product price.

**Functions Used:**

- `SUM()`
- `COUNT()`
- `AVERAGE()`

### 2️⃣ Minimum & Maximum Price

Identified the minimum and maximum product prices using Excel functions.

**Functions Used:**

- `MIN()`
- `MAX()`

### 3️⃣ Logical Function – IF

Created a new column named **Price Range** to categorize products based on their price.

**Business Rule:**

- If Price >= $500 → High Price
- If Price < $500 → Standard Price

**Formula Used:**

```excel
=IF(G3>=500,"High Price","Standard Price")
```

This formula helps categorize products based on a defined price threshold.

### 4️⃣ Conditional Functions – SUMIF & COUNTIF

Performed conditional calculations to extract useful information from the dataset.

#### SUMIF

Calculated the total price of products belonging to the Electronics category.

```excel
=SUMIF(J3:J35,"Electronics",G3:G35)
```

#### COUNTIF

Counted the number of products with a price less than $100.

```excel
=COUNTIF(G3:G35,"<100")
```

These functions helped me understand how to perform category-based and condition-based analysis in Excel.

### 5️⃣ Text Manipulation – LEFT, RIGHT & MID

Extracted useful information from the Product ID column.

New columns created:

| New Column | Excel Function | Purpose |
|---|---|---|
| Day | LEFT() | Extract first 2 characters |
| Country Code | RIGHT() | Extract last 2 characters |
| Month | MID() | Extract characters 4 to 6 |

**Formulas Used:**

#### Day

```excel
=LEFT(A3,2)
```

#### Country Code

```excel
=RIGHT(A3,2)
```

#### Month

```excel
=MID(A3,4,3)
```

These text functions helped transform Product IDs into separate, meaningful fields for analysis.

---

## 📊 Key Results

The following results were calculated from the product dataset.

| Metric | Result |
|---|---:|
| Total Products | 33 |
| Total Price | $9,100 |
| Average Price | $275.76 |
| Minimum Price | $30 |
| Maximum Price | $980 |
| Electronics Category Total | $7,050 |
| Products Priced Below $100 | 11 |

> Note: The total price represents the sum of the product Price column, not price multiplied by quantity.

---

## 🧠 Skills Learned

Through this assignment, I developed practical knowledge of:

### Excel Data Analysis

- Performing basic data exploration.
- Calculating totals, averages, minimums, and maximums.
- Counting records using Excel functions.

### Logical Functions

- Applying IF statements.
- Creating price-based categories.
- Understanding conditional business rules.

### Conditional Aggregation

- Using SUMIF for category-based calculations.
- Using COUNTIF for condition-based counting.

### Text Functions

- Extracting characters from text values.
- Using LEFT, RIGHT, and MID.
- Transforming Product IDs into structured information.

### Data Analyst Foundations

- Understanding structured datasets.
- Applying formulas to solve analytical problems.
- Summarizing data into meaningful insights.
- Preparing datasets for future analysis.

---

## 📁 Project Files

```text
Excel-Data-Exploration/
│
├── Dataset/
│   └── Product Dataset
│
├── Excel Assignment 1 - Data Exploration with answers.xlsx
│
└── README.md
```

---

## 🚀 Learning Outcome

This assignment strengthened my understanding of Microsoft Excel and foundational data analysis techniques.

I learned how to use formulas and functions to explore datasets, categorize information, calculate conditional results, and extract meaningful data from text fields.

This project is one of the initial steps in my journey toward becoming a Data Analyst, and I plan to continue developing my skills in Excel, SQL, Python, Power BI, and other data analytics tools.

---

## 👤 About Me

I am an aspiring Data Analyst, building my technical skills through practical projects and hands-on assignments.

I am passionate about data exploration, problem-solving, and transforming raw data into meaningful insights.

This GitHub repository documents my learning journey and data analytics projects.

---

⭐ Thank you for visiting my repository!
