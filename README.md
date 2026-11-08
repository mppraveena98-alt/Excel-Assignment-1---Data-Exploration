## Excel Assignment 1 - Data Exploration

## 📌 Project Overview

This project demonstrates the practical application of essential Microsoft Excel functions using a product dataset. The assignment focuses on data analysis, conditional calculations, logical functions, and text manipulation.

The objective is to strengthen Excel skills commonly used in data analysis, reporting, and business operations.

## 🛠️ Tools Used

- Microsoft Excel
- Excel Formulas and Functions

## 📊 Tasks Performed

### 1. Basic Statistical Functions – SUM, COUNT & AVERAGE

- Calculated the total price of all products using `SUM()`.
- Counted the number of products using `COUNTA()`.
- Calculated the average product price using `AVERAGE()`.

### 2. Minimum and Maximum – MIN & MAX

- Identified the minimum product price using `MIN()`.
- Identified the maximum product price using `MAX()`.

### 3. Logical Function – IF

Created a new column named **Price Range** to categorize products based on their prices.

- High Price: Products priced greater than or equal to $500.
- Standard Price: Products priced below $500.

**Formula:**

```excel
=IF(Price>=500,"High Price","Standard Price")
```

### 4. Conditional Functions – SUMIF & COUNTIF

- Used `SUMIF()` to calculate the total price of products belonging to the Electronics category.
- Used `COUNTIF()` to count products with a price below $100.

### 5. Text Functions – LEFT, RIGHT & MID

Extracted specific characters from the Product ID column to create three new columns:

- **Day:** Extracted the first two characters using `LEFT()`.
- **Country Code:** Extracted the last two characters using `RIGHT()`.
- **Month:** Extracted the 4th to 6th characters using `MID()`.

**Formulas:**

```excel
=LEFT(A2,2)
=RIGHT(A2,2)
=MID(A2,4,3)
```

## 🎯 Key Learnings

Through this assignment, I gained practical experience in:

- Performing basic statistical analysis in Excel.
- Applying logical conditions using IF statements.
- Performing conditional calculations using SUMIF and COUNTIF.
- Extracting and manipulating text using LEFT, RIGHT, and MID.
- Organizing and analyzing structured product data.

*This project is part of my continuous learning journey in Excel and data analytics.*

