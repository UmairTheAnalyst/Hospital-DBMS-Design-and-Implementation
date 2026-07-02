# 🏥 Hospital Records Management System
### From Spreadsheet to Relational Database — A Design & Implementation Case Study

[![Live Site](https://img.shields.io/badge/Live%20Documentation-Visit%20Site-c04d2a?style=for-the-badge&logo=googlechrome&logoColor=white)](https://umairtheanalyst.github.io/spreadsheet-to-database-migration/)
[![GitHub Repo](https://img.shields.io/badge/GitHub-Project%20Files-1a1917?style=for-the-badge&logo=github&logoColor=white)](https://github.com/UmairTheAnalyst/Hospital-DBMS-Design-and-Implementation/)
[![Built With](https://img.shields.io/badge/Built%20With-MS%20Access%20%7C%20SQL-4a6fc4?style=for-the-badge)]()

---


## Executive Summary

This project documents the complete migration of a hospital patient records system from a
**formula-driven Microsoft Excel workbook** to a **normalized relational database in Microsoft
Access** — and clearly explains *why* that migration was necessary.

The Excel version attempted to track 20 patients across OPD (outpatient) and IPD (inpatient)
care by chaining five `XLOOKUP` formulas across four sheets: one to pull patient names into
billing records, one to calculate each patient's current condition, one to detect OPD-to-IPD
transitions, one to flag that transition on the OPD side, and one to roll up each patient's
total bill back onto their master record. As soon as a patient existed in *both* the OPD and
IPD tables simultaneously — a routine scenario in a real hospital — the formula chain
contradicted itself, silently returning blank cells instead of errors and making billing
accuracy unverifiable.

The core issue was not formula complexity. It was a **structural mismatch**: Excel is a
calculation and analysis tool operating on flat data. It has no primary keys, no foreign key
enforcement, no referential integrity, and no separation between *storing* records and
*querying* them. Every feature needed to manage connected records reliably — unique IDs,
enforced table relationships, cascade rules, typed fields — is architecturally absent from
a spreadsheet.

This project re-architects the same dataset in Microsoft Access using proper relational
database design: five normalized tables, four engine-enforced relationships, dynamic SQL
queries that replace the broken formula chains, and data entry forms that prevent direct
cell editing. The full decision-making process — including every formula that was tried and
exactly why it broke — is documented on the live website above.

---

## Tech Stack & Concepts

| Category | Tools / Concepts |
|---|---|
| **Database Engine** | Microsoft Access 2019+ (.accdb) |
| **Query Language** | Access SQL — `INNER JOIN`, `DateDiff`, `IIf`, `Format`, calculated fields |
| **Database Design** | Relational Database Design, Table Normalization (1NF → 3NF) |
| **Data Integrity** | Primary Keys, Foreign Keys, Referential Integrity, One-to-Many Relationships |
| **Spreadsheet (Legacy)** | Microsoft Excel — `XLOOKUP`, `IF`, `AND`, `IFERROR`, Excel Tables (Structured References) |
| **Documentation** | Custom HTML/CSS static site, Technical Writing, Data Storytelling |
| **Hosting** | GitHub Pages |

---

## Repository Contents

```
Hospital-DBMS-Design-and-Implementation/
│
├── index.html
│   └── Full project documentation site — problem statement, Excel formula
│       breakdown (with annotated screenshots), Access schema, ERD screenshot,
│       SQL query, comparison table, and database overview. Deployed via
│       GitHub Pages at the live link above.
│
├── Hospital_DBMS.accdb
│   └── The working Microsoft Access database. Open in Access to inspect the
│       5-table normalized schema, 4 enforced relationships, saved queries
│       (IPD Billing, OPD Billing, Unpaid IPD Bills, Unpaid OPD Bills),
│       data-entry forms (OPD Visits Form, IPD Admissions Form), and reports.
│
├── Final_Hospital_Patients_record_management.xlsx
│   └── The original Excel workbook, preserved unmodified as a "before"
│       reference. Contains the broken XLOOKUP formula chains that motivated
│       the migration to Access.
│
├── OPD_Billing_-_Formula_bar_of_Patient_Name_column.jpeg
│   └── Screenshot: XLOOKUP pulling patient name from Patients Record into
│       the OPD Billing sheet.
│
├── Current_condition_formula_in_Patient_Records_sheet.jpeg
│   └── Screenshot: Nested IF + XLOOKUP formula determining whether each
│       patient is "Admitted (IPD)", "Active OPD", or "Registered".
│
├── Condition_transition_formula_in_patient_records_sheet.jpeg
│   └── Screenshot: IF + AND + XLOOKUP formula flagging patients present
│       in both OPD and IPD tables as "OPD → IPD Transition".
│
├── current_status_formula_in_OPD_billing_sheet.jpeg
│   └── Screenshot: IF + XLOOKUP formula on the OPD side re-checking
│       whether an OPD patient has since been admitted to IPD.
│
├── Formula_for_fetching_patients_total_amount_from_other_sheets_in_Patient_Records_sheet.jpeg
│   └── Screenshot: IFERROR + IF + XLOOKUP formula attempting to store a
│       single "Total Amount" on the master patient record — the formula
│       that broke silently for OPD-to-IPD transition patients.
│
├── Microsoft_Access_Relationships_between_tables.jpeg
│   └── Screenshot: The actual Access Relationships window showing all four
│       enforced one-to-many connections between the five tables.
│
└── README.md
```

---

## Database Schema

The Access database contains five normalized tables. `Name` exists only in `Patients Record`.
`Age`, `No# of days`, `Total Amount`, `Discount Amount`, and `Net Amount` are stored nowhere —
they are calculated at query time from raw source data.

### Patients Record
| Field | Type | Constraint |
|---|---|---|
| Patient ID | Text | **PK** — unique, engine-enforced |
| Name | Text | |
| DOB | Date | |
| Gender | Text | |
| Registration Date | Date | |
| Mobile No | Text | |
| Alternate Mobile No | Text | |
| Address | Text | |
| Blood Group | Text | |
| Status | Text | |

### OPD Visits
| Field | Type | Constraint |
|---|---|---|
| OPD Visit ID | AutoNumber | **PK** |
| Patient ID | Text | **FK** → Patients Record |
| Visit Date | Date | |
| Bill Type ID | Number | **FK** → OPD Rates |
| Bill type | Text | |
| Quantity | Number | |
| Status | Text | |
| Payment Method | Text | |
| Payment date | Date | |
| Remarks | Memo | |

### IPD Admissions
| Field | Type | Constraint |
|---|---|---|
| Admission ID | AutoNumber | **PK** |
| Patient ID | Text | **FK** → Patients Record |
| Ward ID | Number | **FK** → IPD Rates |
| Ward Name | Text | |
| Bed No# | Number | |
| Admission date | Date | |
| Discharge date | Date | |
| Status | Text | |
| Payment Method | Text | |
| Payment date | Date | |
| Remarks | Memo | |

### IPD Rates
| Field | Type | Constraint |
|---|---|---|
| Ward ID | AutoNumber | **PK** |
| Ward Name | Text | |
| Ward Rate per day | Currency | |
| Room Charges | Currency | |
| Nursing Charges | Currency | |
| Doctor Charges | Currency | |
| Discount Rate | Number | |

### OPD Rates
| Field | Type | Constraint |
|---|---|---|
| Bill Type ID | AutoNumber | **PK** |
| Bill Type Name | Text | |
| Rate | Currency | |
| IsQuantityBased | Yes/No | |
| Discount Rate | Number | |

---

## Enforced Relationships

Four one-to-many relationships are enforced directly by the Access database engine:

| Parent Table | Child Table | Join Key | Cardinality |
|---|---|---|---|
| Patients Record | IPD Admissions | `Patient ID` | 1 → ∞ |
| Patients Record | OPD Visits | `Patient ID` | 1 → ∞ |
| IPD Rates | IPD Admissions | `Ward ID` | 1 → ∞ |
| OPD Rates | OPD Visits | `Bill Type ID` | 1 → ∞ |

A billing record cannot reference a patient, ward, or bill type that does not exist. This
constraint is guaranteed at the engine level and was structurally impossible to replicate
in Excel.

---

## Why Excel Was Not the Right Tool

Six specific capabilities were required that Excel cannot provide:

| # | Gap | Impact in this project |
|---|---|---|
| 1 | **No Primary Key enforcement** | Same Patient ID could be entered twice with no warning |
| 2 | **No referential integrity** | Deleting a patient left billing rows pointing to nobody |
| 3 | **XLOOKUP ≠ a table relationship** | Five stacked formulas contradicted each other when a patient existed in both OPD and IPD tables |
| 4 | **No one-to-many support** | A patient with multiple visits required duplicating their personal details on every row |
| 5 | **No atomic transactions** | An incomplete record (e.g. missing ward type) broke downstream formula totals silently |
| 6 | **Calculated values belong in queries, not master records** | Storing `Total Amount` on the patient record failed the moment a patient had two simultaneous bills |

---

## Key Technical Outcomes

### 1. Third Normal Form (3NF) Normalization

Every repeating attribute was extracted into its own table. Patient demographic data (`Name`,
`DOB`, `Blood Group`, `Mobile No`) is stored exactly once in `Patients Record`. Ward pricing
data (`Ward Rate per day`, `Room Charges`, `Nursing Charges`, `Doctor Charges`,
`Discount Rate`) is stored exactly once in `IPD Rates`. No field in any table depends on
anything other than the primary key of that table.

### 2. Dynamic SQL Replaces Broken Formula Chains

In Excel, `Age`, `No# of days`, `Total Amount`, `Discount Amount`, and `Net Amount` were
either stored as stale calculated fields or silently returned blank for transition patients.
In Access, the `IPD Billing` query calculates all of these live, every time it runs:

```sql
SELECT
  [Patients Record].[Patient ID], [Patients Record].Name,
  [IPD Admissions].[Admission ID], [Patients Record].DOB,

  -- Age calculated from DOB at query time, never stored
  DateDiff("yyyy",[DOB],Date())
    - IIf(Format([DOB],"mmdd") > Format(Date(),"mmdd"), 1, 0) AS Age,

  [Patients Record].Gender, [Patients Record].[Mobile No],
  [IPD Admissions].[Ward ID], [IPD Admissions].[Ward Name],
  [IPD Admissions].[Bed No#],
  [IPD Admissions].[Admission date], [IPD Admissions].[Discharge date],

  -- Duration calculated from dates, never stored
  [Discharge date] - [Admission date] + 1 AS [No# of days],

  [IPD Rates].[Ward Rate per day], [IPD Rates].[Room Charges],
  [IPD Rates].[Nursing Charges], [IPD Rates].[Doctor Charges],

  -- Total Amount calculated from rates, not stored on the patient record
  ([Ward Rate per day] * [No# of days])
    + ([Room Charges] * [No# of days])
    + ([Nursing Charges] * [No# of days])
    + ([Doctor Charges] * [No# of days]) AS [Total Amount],

  [IPD Rates].[Discount Rate],
  [Total Amount] * [Discount Rate]          AS [Discount Amount],
  [Total Amount] - [Discount Amount]        AS [Net Amount],

  [IPD Admissions].Status,
  [IPD Admissions].[Payment Method],
  [IPD Admissions].[Payment date]

FROM [Patients Record]
  INNER JOIN (
    [IPD Admissions]
      INNER JOIN [IPD Rates]
        ON [IPD Admissions].[Ward ID] = [IPD Rates].[Ward ID]
  )
  ON [Patients Record].[Patient ID] = [IPD Admissions].[Patient ID];
```

Updating a ward's daily rate in `IPD Rates` automatically recalculates every patient's bill
across all queries with no manual intervention required.

### 3. The OPD-to-IPD Transition Edge Case

Two patients (P-008 and P-014) transitioned from outpatient to inpatient care. In Excel,
this required three separate formulas just to *label* the situation — a `Current Condition`
formula, a `Condition Transition` formula, and a `Current Status` formula — and even then
their billing totals returned blank because `IFERROR` suppressed the underlying conflict.

In Access, each patient simply has one row in `OPD Visits` and one row in `IPD Admissions`,
both linked via `Patient ID`. The `OPD Billing` query and the `IPD Billing` query each
retrieve their respective records independently with no conflict. No workaround formulas
are needed.

---

## Database at a Glance

| Metric | Value |
|---|---|
| Total Patients | 20 |
| OPD Visit Records | 10 |
| IPD Admission Records | 11 |
| Ward Types (with individual rate structures) | 5 (General, Semi-Private, Private, ICU, CCU) |
| Patients who transitioned OPD → IPD | 2 |
| Saved Queries | 4 (IPD Billing, OPD Billing, Unpaid IPD Bills, Unpaid OPD Bills) |
| Data Entry Forms | 2 (OPD Visits Form, IPD Admissions Form) |
| Enforced Relationships | 4 |

---

## Author

**Umair Asad**
[github.com/UmairTheAnalyst](https://github.com/UmairTheAnalyst)
