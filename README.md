# Python_Expenses_Reconcilition
An automated expense reconciliation system using python and pandas
Python Expenses Reconciliation System

Project Overview

An automated expense reconciliation system built with Python and Pandas.

The system combines expense records from multiple Excel files, compares them with a reference dataset, identifies missing or inconsistent values, corrects discrepancies using reference data, and generates a detailed Excel report.

Features

- Read and combine multiple Excel files.
- Detect exact duplicate records.
- Identify repeated and missing reference numbers.
- Detect missing values.
- Merge expense records with reference data.
- Identify unmatched transactions.
- Clean unnecessary spaces from text values.
- Compare expense data with reference records.
- Automatically correct missing or inconsistent values using reference data.
- Record every correction in a Changes Log.
- Export results into a multi-sheet Excel report.

Technologies Used

- Python
- Pandas
- OpenPyXL
- OS

Input Files

The system uses three Excel files:

1. "expenses_january.xlsx"
2. "expenses_february.xlsx"
3. "transactions_reference.xlsx"

Output Report

The system generates:

"reconciliation_report.xlsx"

The report contains three worksheets:

Worksheet| Description
Corrected Expenses| Expense records after reconciliation
Changes Log| Records old and new values for every correction
Unmatched Expenses| Expense records without a matching reference

How to Run

Install the required libraries:

pip install pandas openpyxl

Run the Python script:

python expenses_reconciliation.py

Make sure the input Excel files are available in the configured folder.

Project Purpose

This project demonstrates practical Python automation skills in Excel data processing, data validation, reconciliation, error identification, automated correction, and report generation.
