# Excel Student Data Analysis Project

## Overview

This project demonstrates the use of Microsoft Excel functions for data analysis and management of student records.

## Dataset

The dataset contains:

* Student ID
* Student Name
* Math Marks
* Science Marks
* English Marks
* Enrollment Date

## Tasks Performed

### 1. Average Calculation

Used the AVERAGE function to calculate the average marks of each student.

Example:
=AVERAGE(C2:E2)

### 2. Grade Assignment

Used IF statements to assign grades based on average marks.

Example:
=IF(G2>=80,"A",IF(G2>=70,"B",IF(G2>=60,"C",IF(G2>=50,"D","F"))))

### 3. Count Students Above 60

Used COUNTIF to count students with average marks greater than 60.

Example:
=COUNTIF(G2:G21,">60")

### 4. VLOOKUP

Used VLOOKUP to retrieve product prices from sales data.

Example:
=VLOOKUP(ProductName,TableRange,ColumnNumber,FALSE)

### 5. XLOOKUP

Used XLOOKUP to dynamically fetch employee salary information.

Example:
=XLOOKUP(EmployeeID,EmployeeRange,SalaryRange)

### 6. Date Analysis

Formatted enrollment dates and extracted date information for analysis.

### 7. FILTER Function

Used FILTER to display students whose average marks are greater than 80.

Example:
=FILTER(B2:G21,G2:G21>80)

## Files Included

* Project-1.xlsx
* README.md
* Screenshots of outputs

## Tools Used

* Microsoft Excel
* GitHub

## Author

Rishit Rajput
