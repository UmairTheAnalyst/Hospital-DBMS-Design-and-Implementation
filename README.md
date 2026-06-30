# 🏥 Hospital Records Management System — From Spreadsheet to Relational Database

[![Live Documentation Site](https://img.shields.io/badge/Live%20Site-View%20Project-c04d2a?style=for-the-badge)](https://umair-asad2001.github.io/spreadsheet-to-database-migration/)
[![Built With](https://img.shields.io/badge/Built%20With-MS%20Access%20%7C%20SQL-4a6fc4?style=for-the-badge)]()

**🔗 Live Documentation Site:** **[umair-asad2001.github.io/spreadsheet-to-database-migration](https://umair-asad2001.github.io/spreadsheet-to-database-migration/)**

---

## Executive Summary

This repository documents the end-to-end migration of a hospital patient records system from a **formula-driven Excel spreadsheet** to a **normalized relational database in Microsoft Access**.

The Excel version relied on chained `XLOOKUP` formulas across four sheets to simulate relationships between patients, OPD visits, and IPD admissions — including a "Current Condition" flag, a "Condition Transition" flag, and a "Total Amount" rollup, all stacked on top of one another. As soon as a single patient existed in more than one table — a routine real-world scenario for a patient moving from outpatient to inpatient care — the formula chain broke silently, returning blank values instead of errors and putting billing accuracy at risk.

This project is a practical case study in **recognizing the limits of spreadsheet-based data management** and **re-architecting the same dataset using core relational database principles**: primary/foreign key enforcement, table normalization, and query-driven calculations. The result is a system where data integrity is guaranteed by the database engine itself, not by the correctness of a formula chain.

---

## Tech Stack & Concepts

| Category | Tools / Concepts Applied |
|---|---|
| **Database Engine** | Microsoft Access (.accdb) |
| **Query Language** | SQL (Access SQL — `INNER JOIN`, `DateDiff`, `IIf`, calculated fields) |
| **Database Design** | Relational Database Design, Table Normalization, Entity Relationship Modeling |
| **Data Integrity** | Primary Keys, Foreign Keys, Referential Integrity, One-to-Many Relationships |
| **Spreadsheet Tooling (Legacy)** | Microsoft Excel — `XLOOKUP`, `IF`, `AND`, `IFERROR`, Excel Tables (Structured References) |
| **Documentation** | HTML/CSS (custom-built static site), Technical Writing, Data Storytelling |

---

## Repository Contents

```text
├── index.html # Full project documentation site (live via GitHub Pages)
├── Hospital_DBMS.accdb # Final relational database — 5 normalized tables, enforced relationships, and SQL queries
├── Final_Hospital_Patients_record_management.xlsx # Original Excel attempt — formulas preserved as-is for reference
├── OPD Billing - Formula bar of Patient Name column.jpeg # XLOOKUP pulling Patient Name into OPD Billing
├── Current condition formula in Patient Records sheet.jpeg # "Current Condition" status-tracking formula
├── Condition transition formula in Patient Records sheet.jpeg # "Condition Transition" OPD→IPD flag formula
├── Current status formula in OPD Billing sheet.jpeg # OPD-side "Current status" cross-check formula
├── Formula for fetching patients total amount...jpeg # The broken "Total Amount" rollup formula
├── Microsoft Access Relationships between tables.jpeg # Actual Access Relationships window (ERD)
└── README.md # You are here
```
| File | Purpose |
|---|---|
| `index.html` | The complete case study — problem framing, Excel formula breakdown (with annotated formula-bar screenshots), Access schema, ERD, comparison table, and SQL queries. Deployed via GitHub Pages. |
| `Hospital_DBMS.accdb` | The working Microsoft Access database. Open in Access to inspect the 5-table schema, the 4 enforced relationships, and saved queries (`IPD Billing`, `OPD Billing`, `Unpaid IPD Bills`, `Unpaid OPD Bills`). |
| `Final_Hospital_Patients_record_management.xlsx` | The original Excel file, kept unmodified as a "before" reference — including the broken `XLOOKUP` chains that motivated the migration. |
| `*.jpeg` screenshots | Annotated formula-bar captures and the Access Relationships window, referenced directly inside `index.html` to ground every claim in real evidence rather than illustrative mockups. |

---

## Key Technical Outcomes

### Eliminated Data Redundancy Through Normalization
The original Excel sheet repeated a patient's `Name`, `DOB`, and `Blood Group` on every billing row. In the Access schema, **`Name` exists in exactly one place** — the `Patients Record` table — and every other table (`OPD Visits`, `IPD Admissions`) references it via `Patient ID` (foreign key) instead of duplicating it.

### Replaced Stored Calculations with Dynamic SQL
In Excel, values like `Age`, `No# of days`, and `Total Amount` were stored as static formula outputs that went stale or returned errors (`#VALUE!`) when source data was incomplete — most notably for patients mid-transition between OPD and IPD. In Access, these are computed **live, on every query run**, directly from source dates and rates:

````sql
DateDiff("yyyy",[DOB],Date())
  - IIf(Format([DOB],"mmdd") > Format(Date(),"mmdd"), 1, 0) AS Age,

[Discharge date] - [Admission date] + 1 AS [No# of days],

([Ward Rate per day] * [No# of days])
  + ([Room Charges] * [No# of days])
  + ([Nursing Charges] * [No# of days])
  + ([Doctor Charges] * [No# of days]) AS [Total Amount]
````

### Enforced Engine-Level Referential Integrity

Four relationships are enforced directly by the Access database engine — not by formulas that can silently fail:

| Parent Table     | Child Table    | Key            | Cardinality |
| ----------------- | -------------- | -------------- | ----------- |
| Patients Record    | IPD Admissions | `Patient ID`   | 1 → ∞       |
| Patients Record    | OPD Visits     | `Patient ID`   | 1 → ∞       |
| IPD Rates           | IPD Admissions | `Ward ID`      | 1 → ∞       |
| OPD Rates           | OPD Visits     | `Bill Type ID` | 1 → ∞       |

This guarantees, at the database level, that a billing record can never reference a patient, ward, or bill type that doesn't exist — a constraint that was structurally impossible to enforce in the spreadsheet version.

### Resolved the Multi-Status Patient Edge Case

The original spreadsheet broke specifically when a patient transitioned from OPD to IPD care, since `XLOOKUP` could only resolve one matching record at a time across two competing sheets — requiring three separate formulas (`Current Condition`, `Condition Transition`, `Current Status`) just to approximate what a relational schema does by default. In the relational model, a patient can simultaneously have related rows in both `OPD Visits` and `IPD Admissions` with zero conflict — each billing event is its own row, linked back to one unambiguous patient record.

---

## Author

**Umair Asad**
[GitHub: @Umair-Asad2001](https://github.com/Umair-Asad2001)

