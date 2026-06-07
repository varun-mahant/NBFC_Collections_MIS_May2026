# NBFC Collections MIS Report

An end-to-end Collections MIS project built in Microsoft Excel to monitor portfolio health, overdue accounts, branch performance, and collection efficiency for an auto loan portfolio.

This project simulates the reporting framework used by NBFCs, banks, and collections teams to track delinquency, recovery performance, and operational efficiency.

---

## Project Overview

The report was built using Excel, Power Pivot, DAX, Pivot Tables, Pivot Charts, and Data Modeling techniques.

The objective was to create a centralized MIS solution capable of:

- Monitoring overdue accounts
- Tracking collection performance
- Measuring branch productivity
- Identifying high-risk branches
- Supporting management decision-making through KPI-driven reporting

---

## Tools & Technologies

- Microsoft Excel
- Power Pivot
- Power Query
- DAX Measures
- Pivot Tables
- Pivot Charts
- Conditional Formatting
- Data Validation
- XLOOKUP
- Data Modeling

---

## Dashboard Modules

### 1. Customer Master

Central repository containing:

- Customer ID
- Customer Name
- Branch Details
- Loan Details
- Disbursement Information
- EMI Information

---

### 2. Collection Log

Tracks collection activities and customer interactions.

Features:

- Agent-wise collections
- Commitment tracking
- Collection amount tracking
- Contact result monitoring
- Collection remarks

---

### 3. Loan Status Tracker

Monitors portfolio health through:

- DPD (Days Past Due)
- Current Loan Status
- Overdue Amount
- Risk Categorization

---

### 4. Overdue & Risk MIS

Portfolio-level monitoring dashboard.

KPIs:

- Total Customers
- Total Loan Amount
- Total Overdue Amount
- Portfolio at Risk (PAR%)
- Accounts at Risk

Analysis:

- DPD Distribution
- Overdue Amount by DPD Bucket
- Aging Summary
- Top Overdue Customers
- Risk Concentration Analysis

---

### 5. Branch Performance Tracker

Branch-wise collections monitoring dashboard.

KPIs:

- Total Branches
- Amount Collected
- Collection Efficiency
- Best Performing Branch
- Highest Overdue Branch

Analysis:

- Collection Efficiency by Branch
- Branch-wise Collection Amount
- Branch Performance Summary
- Overdue Monitoring

---

## Data Model

The project utilizes Excel's Data Model and Power Pivot relationships.

### Tables Used

- Customer Master
- Collection Log
- Loan Status Tracker
- Agent Master

### Relationships

Customer_ID acts as the primary key used to connect multiple reporting tables.

This enables cross-table analysis without extensive lookup formulas.

---

## DAX Measures

### Collection Efficiency

```DAX
Collection Efficiency =
DIVIDE(
    [Sum of Amount_Collected],
    [Sum of Overdue_Amount],
    0
)
