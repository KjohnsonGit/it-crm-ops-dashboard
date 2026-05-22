# IT Support & CRM Operations Dashboard

**Author:** Kenneth Johnson | Business Analyst | Philadelphia, PA
**Tools:** Python · Excel · openpyxl · Power BI-style KPI Design
**GitHub:** [it-crm-ops-dashboard](https://github.com/YourUsername/it-crm-ops-dashboard)

---

## Project Overview

This project is an automated IT Support and CRM Operations analytics workbook built with Python and Excel. It simulates the kind of reporting, root cause analysis, and Agile sprint tracking used in real enterprise IT and business analyst environments — modeled directly on experience supporting a federal civilian agency CRM platform (Salesforce, Jira, Confluence).

The Python script generates a fully formatted, multi-sheet Excel dashboard from scratch — no manual formatting required.

---

## What This Demonstrates

| Skill | How It's Shown |
|---|---|
| Python (openpyxl) | Automated workbook generation via script |
| Excel formulas | Live calculations — no hardcoded values |
| KPI Dashboard Design | Executive summary cards, monthly trend tables |
| Root Cause Analysis | 9-category ticket breakdown with recommendations |
| Agile / Sprint Tracking | Velocity, completion rate, backlog delta by sprint |
| Data Storytelling | Charts, conditional formatting, color-coded metrics |

---

## Dashboard Sheets

### 1. Executive Summary
- 6 KPI cards: Total Tickets, Avg Resolution Time, SLA Compliance, CSAT Score, Open Backlog, Reopen Rate
- Monthly ticket volume and resolution table (12 months, live formulas)
- Bar chart: ticket volume trends
- Line chart: SLA compliance trend

### 2. Root Cause Analysis
- 9 issue categories ranked by volume and escalation rate
- Columns: Ticket Count, % of Total, Avg Resolution Time, Escalation %, Root Cause, Recommended Action
- Conditional formatting: escalation rates flagged red/yellow/green
- Horizontal bar chart: volume by category

### 3. Sprint & Agile Tracker
- 10 sprints tracked: Stories Planned vs Done, Velocity, Completion %, Bugs Found/Resolved, Backlog Delta
- Live completion % formula per sprint
- Line chart: velocity trend across sprints
- Mirrors Jira-style sprint reporting

### 4. README & Methodology
- Project background, data methodology, tool documentation
- Notes on ITIL v4 benchmarks used for synthetic data

---

## How to Run

**Requirements:**
```
pip install openpyxl
```

**Generate the dashboard:**
```
python build_dashboard.py
```

Output: `kenneth_johnson_analyst_dashboard.xlsx`

---

## Files

| File | Description |
|---|---|
| `build_dashboard.py` | Python script that generates the full workbook |
| `kenneth_johnson_analyst_dashboard.xlsx` | Pre-generated Excel dashboard output |
| `README.md` | Project documentation |

---

## Background

This project was built as part of a technical portfolio targeting Business Analyst, IT Analyst, and Data Analyst roles. The data is synthetic, modeled on realistic IT support benchmarks. The dashboard structure reflects real workflows from federal IT consulting, including sprint planning, backlog management, SLA reporting, and root cause documentation.

---

*Kenneth Johnson · kej6032@gmail.com · Philadelphia, PA*
