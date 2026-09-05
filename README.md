# HR Monthly Reporting Automation – Excel VBA

An automated **HR Monthly Management Information System (MIS)** built using **Microsoft Excel and VBA** to streamline recurring HR reporting, data validation, KPI calculation, dashboard updates, department scorecards, and PDF reporting.

## 📊 Project Overview

This project automates a recurring HR monthly reporting cycle.

The solution takes controlled HR source data and processes it through:

- Data cleaning and standardization
- Employee-level validation
- Duplicate detection and removal
- Missing-data flagging
- HR KPI calculation
- PivotTable/PivotChart refresh
- Executive dashboard updates
- Department-level reporting
- PDF report generation

The project uses **synthetic HR data** and is designed as a portfolio demonstration of an automated monthly HR reporting process.

**Reporting Period:** July 2026

---

## 🎯 Business Problem

Monthly HR reporting typically combines multiple datasets such as:

- Employee master data
- Attendance
- Leave
- Performance
- Training

Manual processing of these datasets can create issues such as duplicate employee records, inconsistent dates, missing values, and inconsistent status labels.

This project addresses these challenges by creating a **repeatable and controlled VBA-based reporting workflow**.

---

## 🎯 Project Objectives

- Create a repeatable monthly HR reporting workflow
- Preserve the raw/source workbook
- Create separate clean reporting layers
- Standardize text and date fields
- Validate `Employee_ID` as the employee-level primary key
- Remove duplicate employee records
- Flag missing data
- Calculate standardized HR KPIs
- Refresh PivotTables and PivotCharts
- Update the executive dashboard
- Support department-level scorecards
- Generate department-level PDF reports

---

## 🛠️ Tools & Technologies

- **Microsoft Excel**
- **Excel VBA**
- **PivotTables**
- **PivotCharts**
- **Excel Dashboard**
- **Data Validation**
- **Data Cleaning & Standardization**
- **MIS Reporting**
- **PDF Reporting Automation**

---

## 📁 Source Data Architecture

The project uses multiple HR data sources:

| Sheet | Purpose |
|---|---|
| `Employee_Raw` | Employee master data |
| `Attendance` | Monthly attendance information |
| `Leave` | Leave transactions |
| `Performance` | Employee performance information |
| `Training` | Training records |
| `Department_Master` | Department reference data |
| `Monthly_Targets` | July 2026 department targets |
| `Automation_Guide` | VBA process and KPI requirements |
| `Data_Dictionary` | Field definitions and business rules |

---

## ⚙️ Automation Workflow

```text
HR Source Data
      ↓
Import / Replace
      ↓
Standardize Text & Dates
      ↓
Validate Employee_ID & References
      ↓
Flag Missing Data
      ↓
Remove Duplicate Employees
      ↓
Calculate HR KPIs
      ↓
Refresh PivotTables / PivotCharts
      ↓
Update Executive Dashboard
      ↓
Generate Department PDF Reports
