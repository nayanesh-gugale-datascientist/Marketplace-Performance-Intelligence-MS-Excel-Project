# Excel for Data Analytics — Company-Level Project

# "Marketplace Performance Intelligence"

### Quarterly Business Review (QBR) Dashboard Suite for an Online Marketplace

\---

## STEP 1 OF 10 — Project Idea, Requirements, Working and Final Deliverables

## STEP 2 OF 10 — How to Collect the Data (included below)

\---

## PART 1 — The Project Idea

### 1.1 The one-line idea

Students act as the **Business Analytics team of an online marketplace**. The company's leadership wants to know **how the business is really performing** and **what to fix first**. Using one year-by-year dataset of real orders, students turn raw files into a **clean data model, 9 dashboards, insights, and a storytelling presentation** for management.

### 1.2 Why this project is "company level"

|Feature of real company work|How this project copies it|
|-|-|
|Data sits in **many related tables**, not one neat sheet|9 related files (orders, items, customers, payments, reviews, products, sellers, geolocation, category names)|
|Data is **messy**|Timestamps stored as text, blank delivery dates, Portuguese category names, duplicates in geolocation, one customer with many IDs|
|Work starts with a **business request**, not with data|Students begin from a stakeholder email and KPI definitions|
|Output goes to **decision makers**|Executive dashboard, recommendations, a management presentation|
|Work must be **documented and repeatable**|Data dictionary, cleaning log, KPI sheet, refreshable Power Query steps|
|Numbers must be **trusted**|Reconciliation checks and a QA sign-off before delivery|

### 1.3 The business scenario (the "client")

> \*\*From:\*\* Head of Commerce  \*\*To:\*\* Analytics Team
> \*\*Subject:\*\* Quarterly Business Review — need a dashboard pack by month-end
>
> "We sell through a marketplace model: customers order, independent sellers ship, and we handle payments and customer experience. Leadership meets next month and I need to answer five things: \*\*(1)\*\* Is revenue growing, and where does it come from? \*\*(2)\*\* Which products and categories matter most? \*\*(3)\*\* Who are our customers and do they come back? \*\*(4)\*\* Are we delivering on time, and does late delivery hurt satisfaction? \*\*(5)\*\* Which sellers and states need attention? Please give me dashboards, the key insights, and a short story I can present, with clear recommendations."

### 1.4 The business questions the project will answer (grouped by dashboard)

1. **Sales \& Revenue:** How much revenue per month? Growth month-over-month? Best weekday/hour? Average order value (AOV)?
2. **Products \& Categories:** Top 10 categories by revenue and by orders? Price vs freight cost? Which categories are growing or shrinking?
3. **Customers:** How many unique customers? New vs repeat? Who are the best customers (RFM segments)?
4. **Delivery \& Logistics:** Average delivery days? On-time vs late %? Which states are slowest? How much late delivery per month?
5. **Customer Satisfaction:** Average review score? Share of 1-star reviews? Does late delivery lower the score?
6. **Sellers:** Top sellers by revenue? Sellers with the most late orders or worst reviews? How concentrated is revenue?
7. **Geography:** Revenue and orders by state and city? Where do customers and sellers sit?
8. **Payments:** Which payment types are most used? How many installments? Credit-card installment behavior by order value?
9. **Executive Summary:** The 8–10 KPIs a CEO checks in 30 seconds.

\---

## PART 2 — The Requirements

### 2.1 Software

|Item|Requirement|Why|
|-|-|-|
|Microsoft Excel|**Microsoft 365** or Excel 2019/2021 **for Windows**|Needs **Power Query** and **Power Pivot (Data Model)**; Mac Excel has limited support for these|
|Browser + free **Kaggle** account|Any|To download the dataset|
|PDF reader, PowerPoint (or Google Slides)|Any|Final presentation|
|(Optional) MySQL + Workbench|Already installed|Practice loading the data into a database before Excel|
|(Optional) Python|Already known|Extra checks and advanced analysis|

### 2.2 Computer

* **8 GB RAM minimum, 16 GB recommended.** The data has about 100,000 orders and 112,000 order lines. It fits in Excel's Data Model but is slow on low-memory PCs.
* About 1 GB free disk space (raw files, cleaned workbook, exports).

### 2.3 Skills students need (all already covered in your course)

Formulas (IF, XLOOKUP, SUMIFS, TEXT, date functions) · Tables · PivotTables · Charts · Conditional formatting · Power Query basics · Data Model and relationships · basic DAX measures · slicers.

### 2.4 Team roles (to copy real companies)

|Role|What the student does|Real-world job|
|-|-|-|
|**Business Analyst**|Writes requirements, KPI definitions, final story|BA / Product Analyst|
|**Data Analyst**|Cleans data, builds measures, finds insights|Data Analyst|
|**BI Developer**|Builds the dashboards and layout|BI Developer|
|**QA / Reviewer**|Checks every number against the source|Data QA / Reviewer|
|A team of 4 can rotate roles; a solo student wears all four hats.|||

### 2.5 Folder structure (set up on day 1)

```
Marketplace\_Performance\_Intelligence/
├── 01\_Requirements/      (charter, KPI definitions)
├── 02\_Raw\_Data/          (the 9 CSV files, NEVER edited)
├── 03\_Data\_Documentation/(intake log, data dictionary, cleaning log)
├── 04\_Working\_Files/     (Power Query / workbook drafts)
├── 05\_Final\_Workbook/    (final Excel dashboard file)
├── 06\_Reports\_and\_Slides/(insights report, PPT, one-page summary)
└── 07\_QA\_and\_Handover/   (QA checklist, user guide)
```

**Rule #1 in real companies: the raw data folder is never edited.** All cleaning happens in copies or in Power Query steps.

**File naming:** `ProjectName\_Document\_v1.0\_YYYY-MM-DD` (e.g., `MPI\_Dashboard\_v1.0\_2026-10-19.xlsx`).

\---

## PART 3 — STEP 2: How to Collect the Data

### 3.1 The dataset (recommended)

**Brazilian E-Commerce Public Dataset by Olist** — real commercial data published by Olist, a Brazilian marketplace company. It holds information on about **100,000 orders from 2016 to 2018** from multiple marketplaces in Brazil, covering order status, price, payment, freight, customer location, product attributes and customer reviews.

* **Where:** Kaggle → `https://www.kaggle.com/datasets/olistbr/brazilian-ecommerce`
* **Cost:** free (Kaggle account needed)
* **Format:** 9 CSV files
* **Licence:** open the dataset's Kaggle page and read the licence/usage terms before use; it is published for learning and analysis. Note the source and licence in your Data Intake Log.
* **Currency:** all amounts are in **Brazilian Real (R$)**. We will keep R$ and not convert.

### 3.2 How to download (exact steps)

1. Go to **kaggle.com** and create a free account (or sign in)
2. Open the dataset link above
3. Click **Download** → **Download dataset as zip**
4. Create the project folder structure from 2.5
5. Extract the ZIP into `02\_Raw\_Data/`
6. Confirm you see **9 files**:

|#|File name|What it holds|Approx. rows|
|-|-|-|-|
|1|`olist\_orders\_dataset.csv`|**Core table**: one row per order, status and all timestamps|99,441|
|2|`olist\_order\_items\_dataset.csv`|Products inside each order, seller, **price, freight**|112,650|
|3|`olist\_customers\_dataset.csv`|Customer IDs and location|99,441|
|4|`olist\_order\_payments\_dataset.csv`|Payment type, installments, payment value|\~104,000|
|5|`olist\_order\_reviews\_dataset.csv`|Review score and comments|\~99,000|
|6|`olist\_products\_dataset.csv`|Product category, size, weight|\~33,000|
|7|`olist\_sellers\_dataset.csv`|Seller location|\~3,100|
|8|`olist\_geolocation\_dataset.csv`|Zip code to latitude/longitude|\~1,000,000|
|9|`product\_category\_name\_translation.csv`|Portuguese to English category names|71|

(Rows for #1–#3 come from published row counts of the dataset; the others are approximate, so record your own counts in the intake log.)

### 3.3 Important warnings before you open the files

1. **Do not double-click the CSVs to open them in Excel.** Excel may break Portuguese accents (é, ã, ç) and convert long IDs or ZIP codes. Always use **Data → Get Data → From Text/CSV**, set File Origin to **UTF-8 (65001)**, then **Transform Data**.
2. **Do not load `olist\_geolocation\_dataset.csv` into a worksheet.** It has around one million rows. In this project we will either skip it or summarize it in Power Query first.
3. **Never edit the raw files.** Save the originals untouched.

### 3.4 Data Intake Log (fill this in on day 1)

Real companies record where data came from before using it.

|Field|Your entry|
|-|-|
|Dataset name|Brazilian E-Commerce Public Dataset by Olist|
|Source URL|kaggle.com/datasets/olistbr/brazilian-ecommerce|
|Downloaded on|(date)|
|Downloaded by|(name)|
|Licence / terms|(copy from the Kaggle page)|
|Number of files|9|
|Row count per file|(open each file in Power Query and note it)|
|Known limits|Brazil only, 2016–2018, historical data|

### 3.5 Data quirks to expect (we will handle every one of these)

|Quirk|Why it matters|
|-|-|
|`customer\_id` is created **per order**; `customer\_unique\_id` identifies the real person|Counting `customer\_id` makes almost everyone look like a one-time buyer. Repeat-customer analysis must use `customer\_unique\_id`.|
|All dates are **text** like `2017-10-02 10:56:33`|Must be converted to Date/Time for month, weekday, delivery-days work|
|Some orders have blank delivery dates|Canceled/unavailable/in-transit orders; must be filtered correctly for delivery KPIs|
|Category names are in **Portuguese**|Needs the translation table|
|One order can have **many items and many payments**|Joining tables carelessly **multiplies revenue** (the most common real-world analytics error)|
|Some reviews have no comment|Blank is normal, not an error|
|Column names contain a spelling slip (`product\_name\_lenght`)|A real example of messy source systems|
|Sparse data at the very start (2016) and the end (2018)|Needs care when comparing months or years|

### 3.6 Backup dataset options (only if Kaggle is blocked)

|Option|Notes|
|-|-|
|UCI **Online Retail II**|Real UK retailer transactions in one table; simpler (no seller or delivery data), so fewer dashboards possible|
|Our own generated practice data|I can create a realistic synthetic marketplace dataset if you prefer|
|Tell me if you want either; otherwise we proceed with Olist.||

## PART 4 — The Whole Working of the Project (How a real company project runs)

|Phase|Name|What happens|Output (what you produce)|Real-world mapping|
|-|-|-|-|-|
|**0**|Project kickoff|Read the stakeholder request, agree scope, set up folders|Project charter, folder structure|Ticket / kickoff meeting|
|**1**|Requirement gathering|Turn questions into KPIs with exact definitions|KPI definition sheet|BRD / KPI dictionary|
|**2**|Data collection|Download data, log the source|Raw data folder, intake log|Data request / access|
|**3**|Data understanding (audit)|Row counts, columns, types, blanks, duplicates, relationships|Data dictionary, data quality report|Data profiling|
|**4**|Data cleaning|Fix types, text, blanks, duplicates, outliers; document every change|Clean tables, cleaning log|ETL / data prep|
|**5**|Data transformation \& modeling|Join tables, add calculated columns, build relationships in the Data Model|Star-schema Data Model|Data modeling|
|**6**|Data analysis|Write DAX measures, compare periods, segment customers (RFM), test ideas|Measure library, analysis notes|Analytics|
|**7**|Dashboard building|Build 9 dashboards with slicers and consistent design|Dashboard workbook|BI development|
|**8**|Insight generation|Turn charts into findings, causes and recommendations|Insights log|Analysis write-up|
|**9**|Storytelling|Arrange insights into a story for management|Presentation deck, executive summary|Business review|
|**10**|QA and handover|Reconcile numbers, test refresh, write the user guide, get sign-off|QA checklist, user guide, final package|UAT / go-live / handover|

**Good-practice habits we will follow (these separate students from professionals):**

* Keep a **cleaning log**: what we changed, how many rows, why
* **Reconcile**: totals in the model must match the totals in the raw files
* One **KPI definition** for each metric, used everywhere
* Never hard-code numbers in dashboards; everything comes from the model so it **refreshes**
* Separate sheets for **data, calculations, and dashboards**

\---

## PART 5 — The Plan: What We Will Build

### 5.1 The 9 dashboards

|#|Dashboard|Main visuals (planned)|
|-|-|-|
|1|**Executive Summary**|KPI cards (Revenue, Orders, AOV, Customers, On-time %, Avg Review), monthly revenue trend, top 5 categories, top 5 states|
|2|**Sales \& Revenue**|Revenue by month and year, growth %, weekday and hour heat map, AOV trend|
|3|**Product \& Category**|Top 10 categories, revenue vs orders, price distribution, freight share of price|
|4|**Customer Analysis**|New vs repeat, orders per customer, RFM segments, customers by state|
|5|**Delivery \& Logistics**|Avg delivery days, on-time %, late orders by month, slowest states, estimated vs actual|
|6|**Customer Satisfaction**|Review score distribution, 1-star %, score by delivery status, score by category|
|7|**Seller Performance**|Top sellers, revenue concentration (Pareto), late-order rate, seller ratings|
|8|**Geography**|Revenue and orders by state and city, customer-to-seller flows|
|9|**Payments**|Payment type mix, installment distribution, order value by payment type|

### 5.2 KPI definitions (agreed up front)

|KPI|Definition|
|-|-|
|**Product Revenue**|Sum of `price` for items in **delivered** orders (excludes freight)|
|**Freight Revenue**|Sum of `freight\_value` for the same items|
|**GMV**|Product Revenue + Freight Revenue|
|**Orders**|Count of distinct `order\_id`|
|**AOV**|Product Revenue ÷ Orders|
|**Unique Customers**|Count of distinct `customer\_unique\_id`|
|**Repeat Customer %**|Customers with 2+ orders ÷ unique customers|
|**Delivery Days**|`order\_delivered\_customer\_date` − `order\_purchase\_timestamp` (delivered orders only)|
|**On-Time %**|Orders delivered on or before `order\_estimated\_delivery\_date` ÷ delivered orders|
|**Avg Review Score**|Mean of `review\_score` (1–5)|
|**Cancellation Rate**|Canceled orders ÷ all orders|
|**Freight % of Price**|Freight Revenue ÷ Product Revenue|
|(We will confirm or adjust these with you when we see the data.)||

### 5.3 The technical flow

```
CSV files → Power Query (clean) → Excel Data Model (relationships)
         → DAX measures → PivotTables / Charts → Dashboards → Story
```

This is the same flow you learned in the Data Pipeline lesson.

### 5.4 How we will divide the work (important, please read)

* **I will do the heavy build** once you share the data: data audit, cleaning, joins, KPI calculations, analysis, insights and the storytelling content.
* I build the workbook with code, so the finished file will contain **cleaned data, KPI tables and native Excel charts**.
* **Interactive pieces that exist only inside Excel** (Power Query refresh steps, Data Model relationships, DAX measures, PivotTables and slicers) cannot be created by me directly. For those I will give students **exact, copy-paste step-by-step instructions plus the Power Query (M) and DAX code**, so students rebuild the interactive layer themselves inside Excel. That is also the best way for them to learn.

\---

## PART 6 — What We Will Hold in Hand When the Project is Done

### 6.1 Final deliverables

|#|Deliverable|Format|What it is|
|-|-|-|-|
|1|**Project Charter**|MD / Word|Scope, questions, roles, timeline (this document)|
|2|**KPI Definition Sheet**|Excel|Exact formula for every KPI|
|3|**Data Intake Log + Data Dictionary**|Excel|Source, tables, columns, meaning, types|
|4|**Data Quality Report + Cleaning Log**|Excel / Doc|Every issue found and how it was fixed, with row counts|
|5|**Clean Data Model Workbook**|Excel (.xlsx)|Cleaned tables, relationships, measures|
|6|**Dashboard Workbook**|Excel (.xlsx)|The 9 dashboards|
|7|**Power Query + DAX Script Pack**|Text / MD|Reusable code and step-by-step rebuild guide|
|8|**Insights and Recommendations Report**|Doc|Findings, causes, recommended actions|
|9|**Management Presentation**|PowerPoint|12–15 slides telling the story|
|10|**One-Page Executive Summary**|PDF / Doc|The whole project on one page|
|11|**QA Checklist + Reconciliation Report**|Excel|Proof the numbers are right|
|12|**User Guide / Handover Note**|Doc|How to refresh and use the dashboards|
|13|**Student Portfolio Pack**|Doc|GitHub-ready README, LinkedIn post text, resume bullet points, interview questions and answers|

### 6.2 Definition of "Done" (success checklist)

* \[ ] All 8 data files loaded; row counts match the intake log
* \[ ] Every cleaning step documented in the cleaning log
* \[ ] Revenue in the model **equals** the raw-file total (no duplication from joins)
* \[ ] All KPIs follow the written definitions
* \[ ] 9 dashboards built; slicers work; each page answers its business question
* \[ ] At least 15 written insights, each with evidence and a recommendation
* \[ ] Presentation tells a clear story: situation → problem → evidence → action
* \[ ] A second person (QA role) has checked the numbers
* \[ ] The workbook refreshes without errors
* \[ ] Handover guide and portfolio pack complete

### 6.3 What students gain

* A **portfolio-ready, company-style project** (not a toy exercise)
* Practical use of Power Query, Data Model, DAX, PivotTables, dashboards
* The habit of **KPI definitions, documentation, QA, and storytelling**
* Interview-ready answers to "Tell me about a project you did end to end"

\---

## PART 7 — Roadmap: Step by Step

|Step|What happens|Who|Status|
|-|-|-|-|
|**1**|Project idea, requirements, working, deliverables|Me|✅ This document|
|**2**|Collect data (download, folder setup, intake log)|You|⬅ **You are here**|
|**3**|Data audit and data dictionary|Me|Next, after you share the files|
|**4**|Data cleaning (with cleaning log)|Me + students|Pending|
|**5**|Data transformation and model design|Me + students|Pending|
|**6**|KPI measures and data analysis (RFM, cohorts, delivery impact)|Me + students|Pending|
|**7**|Building the 9 dashboards|Me + students|Pending|
|**8**|Insights and recommendations|Me|Pending|
|**9**|Storytelling: deck, report, executive summary|Me|Pending|
|**10**|QA, handover, portfolio pack|Me + students|Pending|

### Suggested timeline for students (4-week sprint)

|Week|Phases|Goal|
|-|-|-|
|Week 1|0–3|Kickoff, requirements, data collected, data understood|
|Week 2|4–5|Clean data and a working Data Model|
|Week 3|6–7|KPIs, analysis, dashboards|
|Week 4|8–10|Insights, story, QA, final presentation|



