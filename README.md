<div align="center">

# 📊 MedVault Inventory Management & Stock Monitoring Dashboard

![Dashboard Preview](./MedVault%20Dashboard%20Preview/MedVault-Inventory%20Dashboard%20Screenshot.jpg)

_Turning raw inventory transactions into real-time, decision-ready visibility._

</div>

---

## 📌 Project Overview

MedVault is a fully automated inventory management and stock monitoring system built entirely in Microsoft Excel, designed to simulate the operational backbone of a multi-department retail and pharmaceutical supply business.

The project replaces fragmented, manual stock tracking (notebooks, disconnected spreadsheets, and tribal knowledge) with a single, centralized workbook that captures every inventory transaction, calculates live stock balances, flags re-order risks automatically, and visualizes the entire operation through an interactive, slicer-driven dashboard.

It was designed to demonstrate that advanced inventory intelligence, the kind typically associated with paid ERP or POS software, can be engineered using Excel's native toolset alone: dynamic formulas, structured tables, Pivot Tables, Pivot Charts, slicers, and conditional logic.

---

## 🧭 Executive Summary

|                        |                                                                                                        |
| ---------------------- | ------------------------------------------------------------------------------------------------------ |
| **Project Type**       | Inventory Management & Business Intelligence Dashboard                                                 |
| **Industry Simulated** | Pharmaceutical Retail / Multi-Department Supply Store                                                  |
| **Core Tool**          | Microsoft Excel (Tables, Pivot Tables, Pivot Charts, Slicers, Dynamic Formulas, VBA macro integration) |
| **Catalogue Size**     | 85 SKUs across 5 departments                                                                           |
| **Automation Level**   | Auto-fill data entry, live stock recalculation, automated re-order alerts                              |
| **Reporting Layer**    | Interactive one-page dashboard with cross-filtering slicers                                            |
| **Outcome**            | A reusable, scalable inventory framework adaptable to any retail, healthcare, or distribution business |

---

## 🎯 Business Problem

Most small and mid-sized businesses manage inventory reactively rather than proactively. Common pain points this project was built to solve include:

- **Delayed visibility** — stock figures are often "a few days behind," discovered only when a customer asks for an item that has quietly run out.
- **Manual, error-prone data entry** — staff retyping item names and prices for every transaction, increasing the risk of mismatched records.
- **No early warning system** — reordering happens too late (stockouts) or too early (overstocking, tied-up capital), because nobody has a live reorder threshold view.
- **Fragmented reporting** — sales, stock levels, and supplier performance often live in separate, disconnected sheets with no unified view.
- **Lack of accessible business intelligence** — dashboard-style reporting is usually gated behind paid software subscriptions, putting it out of reach for smaller operations.

The objective was to solve all five problems using a tool nearly every business already owns.

---

## 🏗️ Solution Architecture

The workbook is structured as a relational system of six interconnected sheets, each serving a distinct role in the data pipeline:

```
┌─────────────────┐
│   Item List      │  Master catalogue: GL Codes, Categories, Departments,
│  (Master Data)    │  Units, Prices, Re-Order Levels
└────────┬─────────┘
         │ VLOOKUP
         ▼
┌─────────────────┐
│  Daily Records    │  Transaction ledger: GL Code dropdown auto-fills
│ (Data Entry Layer) │  Item Name, Department, and Unit Price; QTY entry
└────────┬─────────┘  auto-calculates Line Value
         │ SUMIFS
         ▼
┌─────────────────┐
│ Stock Level &     │  Live stock engine: Opening Stock + Inflow − Outflow
│    Status          │  = Current Stock, with automated RAG status flags
└────────┬─────────┘
         │
         ▼
┌─────────────────┐
│  Re-Order Level   │  Threshold monitoring: variance analysis and
│    Register         │  prioritized restocking actions
└────────┬─────────┘
         │
         ▼
┌─────────────────┐
│ Dashboard Report  │  KPI cards, monthly trend analysis, top movers,
│  & Dashboard      │  supplier performance, and live stock alerts
└─────────────────┘
```

**Data flow principle:** a single entry point (selecting a GL Code in Daily Records) cascades automatically through lookup and aggregation formulas to update every downstream sheet, with zero duplicate data entry.

---

## 🖥️ Dashboard Features

The final dashboard consolidates the entire inventory operation into one interactive view:

- **Filter Panel** — Department and Month slicers that cross-filter every chart and KPI simultaneously
- **KPI Strip** — Total Stock Value, Total SKUs, In Stock, Low Stock, Out of Stock, and Reorder Required counts
- **Insight Summary** — auto-generated, plain-language observations (e.g. _"96% of items are currently available"_) that translate raw numbers into instant business takeaways
- **Inventory Health Indicator** — a single-glance status label (Healthy / Needs Attention) driven by live stock conditions
- **Value by Department** — donut chart breaking down stock value contribution across departments
- **Stock Status Breakdown** — donut chart visualizing In Stock vs Low Stock vs Out of Stock proportions
- **Monthly Inflow vs Outflow** — clustered column chart tracking stock movement trends across the reporting period
- **Top 5 Outflow Items** — horizontal bar chart identifying the fastest-moving products
- **Top Suppliers** — horizontal bar chart ranking suppliers by delivered quantity
- **Value by Category** — ranked list view surfacing where capital is concentrated across product categories
- **Automated Refresh** — VBA-driven refresh logic ensuring the dashboard reflects the latest entries without manual intervention

---

## 📐 KPI Definitions

| KPI                    | Definition                                            | Formula Logic                              |
| ---------------------- | ----------------------------------------------------- | ------------------------------------------ |
| **Total Stock Value**  | Current inventory value at cost                       | `SUMPRODUCT` of Current Stock × Unit Price |
| **Total SKUs**         | Total number of distinct items in the catalogue       | `COUNTA` of GL Code column                 |
| **In Stock**           | Items with current stock above the re-order threshold | `COUNTIF` Status = "In Stock"              |
| **Low Stock**          | Items at or below re-order level but not yet at zero  | `COUNTIF` Status = "Low Stock"             |
| **Out of Stock**       | Items with zero current stock                         | `COUNTIF` Status = "Out of Stock"          |
| **Reorder Required**   | Items flagged for immediate restocking action         | `COUNTIF` Action = "Order Now" / "Urgent"  |
| **Current Stock**      | Real-time on-hand quantity per item                   | Opening Stock + Inflow − Outflow           |
| **Inventory Turnover** | Rate at which stock is sold and replaced              | Total Outflow ÷ Average Stock Held         |

---

## 🛠️ Technical Stack

| Component                            | Purpose                                                                          |
| ------------------------------------ | -------------------------------------------------------------------------------- |
| **Microsoft Excel Tables**           | Structured, auto-expanding data ranges for all transactional and master data     |
| **VLOOKUP / IFERROR**                | Auto-fill logic for Item Name, Department, and Unit Price upon GL Code selection |
| **SUMIFS / COUNTIFS**                | Real-time aggregation of inflow, outflow, and status counts                      |
| **Data Validation (Dropdown Lists)** | Controlled data entry for GL Code, Type, Department, Gender, and Month fields    |
| **Conditional Formatting**           | RAG (Red/Amber/Green) visual flags across stock status and reorder actions       |
| **Pivot Tables & Pivot Charts**      | Aggregated, slicer-ready reporting layer for dashboard visuals                   |
| **Slicers & Report Connections**     | Cross-filtering interactivity linking multiple Pivot Tables simultaneously       |
| **Named Ranges**                     | Centralized, maintainable reference lists for dropdown validation                |
| **VBA (Worksheet Change Event)**     | Automated `RefreshAll` trigger to keep dashboard data current                    |
| **Custom Number Formatting**         | Currency formatting (₦) and visual consistency across all reporting sheets       |

---

## 📈 Business Impact

The system is designed to mirror real operational value:

- **Eliminates manual lookup errors** — auto-fill logic on data entry removes the need to retype item names or prices, reducing transcription mistakes to near zero
- **Reduces stockout risk** — automated re-order alerts surface at-risk items before they become unavailable, rather than after a customer complaint
- **Improves decision speed** — what previously required manually cross-referencing multiple sheets is now answered in one filtered glance via the dashboard
- **Increases inventory accountability** — supplier and department-level breakdowns make it possible to track exactly where stock value is concentrated and who is supplying it
- **Removes software cost barriers** — delivers BI-grade reporting capability without requiring a paid inventory management platform, making it viable for businesses with limited tooling budgets

---

## 🧠 Skills Demonstrated

- Advanced Excel formula engineering (nested `VLOOKUP`, `SUMIFS`, `IFERROR`, `SUMPRODUCT`, `INDEX`/`MATCH`)
- Relational data modeling across multiple interconnected sheets
- Dashboard design and information hierarchy (KPI placement, visual flow, filter logic)
- Data validation and controlled-entry system design
- Conditional formatting for status-driven visual communication
- Pivot Table and Pivot Chart construction with multi-table slicer connectivity
- Process automation thinking (minimizing manual touchpoints in a recurring workflow)
- Basic VBA scripting for refresh automation
- Business analysis: translating operational pain points into measurable KPIs
- Visual storytelling: converting a working spreadsheet into a stakeholder-ready report

---

## 🚀 Future Enhancements

- [ ] Migrate the data model to Power Query / Power Pivot for native multi-table relationships at scale
- [ ] Introduce a historical stock snapshot mechanism to enable true point-in-time ("as of [month]") inventory views
- [ ] Add a rolling date-range slicer for more granular period-over-period analysis
- [ ] Build a stock value trend line chart once multiple periods of real transactional data are available
- [ ] Extend the system into Power BI for cloud-based, multi-user dashboard access
- [ ] Introduce supplier lead-time tracking to forecast optimal reorder timing
- [ ] Add barcode/QR-code-based data entry to further reduce manual input

---

## 👤 Author

**Kehinde Ojo**
Data Analyst | Business Intelligence & Financial Analytics | Process Automation

🌐 [Portfolio](https://kehindeojo.netlify.app)
💼 [LinkedIn](https://www.linkedin.com/in/kehindeojo-analyst)
💻 [GitHub](https://github.com/KehindeOjo)

<div align="center">

_If you found this project useful or interesting, consider starring ⭐ the repository._

</div>
