# 🏏 ICC Cricket World Cup 2023 — Excel Analytics Dashboard

An interactive Microsoft Excel dashboard built to explore and analyse ICC Men's Cricket World Cup 2023 match data. The project turns raw match records into visual summaries of team performance, batting and bowling statistics, toss decisions, venues, and match outcomes.

> **Project type:** Data Analytics / Excel Dashboard  
> **Primary tool:** Microsoft Excel  
> **Data source:** `Raw_Data` worksheet in the workbook  
> **Workbook:** `CWC2023.xlsx` *(use the exact filename committed to this repository)*

---

## 📌 Table of Contents

- [Project Overview](#-project-overview)
- [Project Objectives](#-project-objectives)
- [Dashboard Features](#-dashboard-features)
- [Key Performance Indicators](#-key-performance-indicators)
- [Analysis Areas](#-analysis-areas)
- [Tools and Excel Features](#-tools-and-excel-features)
- [Workbook Structure](#-workbook-structure)
- [How to Use the Dashboard](#-how-to-use-the-dashboard)
- [Methodology and Metric Definitions](#-methodology-and-metric-definitions)
- [Data Quality and Validation](#-data-quality-and-validation)
- [Repository Structure](#-repository-structure)
- [Screenshots](#-screenshots)
- [How to Run the Project](#-how-to-run-the-project)
- [Findings and Interpretation](#-findings-and-interpretation)
- [Limitations](#-limitations)
- [Future Improvements](#-future-improvements)
- [Author](#-author)
- [Disclaimer](#-disclaimer)

---

## 📊 Project Overview

The purpose of this project is to demonstrate how Microsoft Excel can be used to organise, analyse, and visualise sports data. The dashboard provides a single place to explore important statistics from the 2023 Cricket World Cup.

Instead of manually reviewing rows of match data, users can view KPIs, compare teams, investigate toss decisions, examine venue patterns, and review match results through charts and filters.

The workbook includes the original match-data worksheet, a dashboard worksheet, and a tournament-highlights worksheet. Worksheet names may vary slightly depending on the final version of the workbook.

## 🎯 Project Objectives

- Transform raw cricket match data into a structured format suitable for analysis.
- Summarise important match and team performance indicators.
- Compare teams using wins, scores, run rates, wickets, and boundaries where the source fields support those measures.
- Explore toss decisions and whether the toss-winning team also won the match.
- Compare match counts and scoring patterns across cities and stadiums.
- Present notable tournament achievements in a dedicated highlights section.
- Practise Excel data analysis, formula use, data visualisation, and dashboard design.

## ✨ Dashboard Features

- **KPI cards:** Quick view of key match and performance statistics.
- **Team performance analysis:** Compare match wins and other available team metrics.
- **Batting and scoring analysis:** Explore scores, average scores, run rates, fours, sixes, and boundaries when available.
- **Bowling analysis:** Review wicket-related statistics available in the source data.
- **Toss analysis:** Compare toss decisions and match outcomes.
- **Venue analysis:** Explore matches and scoring patterns by city or stadium.
- **Match-result analysis:** Examine winning margins and result categories when supported by the dataset.
- **Interactive filtering:** Use connected slicers or other filters provided in the workbook to narrow the displayed data.
- **Tournament highlights:** A separate summary of notable tournament achievements.

> The exact KPIs, charts, and available filters depend on the final workbook version and the fields in the source data.

## 📈 Key Performance Indicators

The dashboard may include the following indicators, depending on the final workbook configuration:

| KPI | Description |
|---|---|
| Total Matches | Number of unique matches represented in the selected data. |
| Total Teams | Number of distinct teams represented in the data. |
| Total Venues | Number of distinct venues represented in the data. |
| Total Runs | Total runs, based on the score fields and calculation definition used. |
| Average Run Rate | Average run rate calculated according to the workbook's chosen method. |
| Total Wickets | Wicket total according to the source fields and aggregation method. |
| Highest Team Score | Highest team-innings score found in the selected data. |
| Most Successful Team | Team with the most wins within the scope of the calculation. |

**Important:** KPI definitions should be checked against the workbook formulas. For example, avoid counting a match twice when the source has separate rows for each team, and do not add together match-level totals that already include both teams.

## 🔎 Analysis Areas

### 1. Team Performance
Compare teams using match wins, matches played, and winning percentage where the required fields are available. Team comparisons should use consistent definitions and the same tournament scope.

### 2. Batting and Scoring
Explore high scores, average scores, run rates, fours, sixes, and total boundaries. Use the source fields carefully to distinguish team-innings totals from combined match totals.

### 3. Bowling Performance
Review available wicket statistics and related match indicators. Interpret wicket counts according to how they are stored in the source data.

### 4. Toss Analysis
Explore batting-first and fielding-first decisions and compare the toss winner with the eventual match winner. This describes patterns in the tournament data; it does not prove that winning the toss causes a match victory.

### 5. Venue Analysis
Compare matches hosted by cities or stadiums and investigate scoring patterns across venues. Consider that venues may host different numbers of matches.

### 6. Match Results
Review match outcomes and winning margins, such as wins by runs or wickets, where these fields are included and consistently classified.

### 7. Tournament Highlights
The highlights worksheet summarises notable achievements from the tournament, including:
- **Champion:** Australia
- **Most runs:** Virat Kohli — 765 runs
- **Most wickets:** Mohammed Shami — 24 wickets
- **Highest individual innings:** Glenn Maxwell — 201 not out against Afghanistan
- **India's league-stage record:** 9 wins from 9 matches

These are tournament-wide facts, not necessarily values that change when dashboard slicers are selected.

## 🧰 Tools and Excel Features

- Microsoft Excel
- Excel Tables and structured references, where used
- Excel formulas and functions
- PivotTables and PivotCharts, where used
- Slicers and date filters/timelines, where available
- Conditional formatting
- Charts and KPI cards
- Data cleaning and validation

Only list a feature as implemented if it is actually present in the committed workbook.

## 📁 Workbook Structure

The workbook may contain the following worksheets:

| Worksheet | Purpose |
|---|---|
| `Raw_Data` or `CWC2023` | Source match records used for analysis. |
| `Dashboard` | Main interactive dashboard with KPIs, charts, and filters. |
| `Tournament Highlights` | Summary of notable tournament achievements. |
| Other analysis worksheets | Supporting summaries or calculations, if included in the final workbook. |

**Single-source-of-truth note:** If the final dashboard is intended to use only the Raw Data worksheet as its source, document that arrangement accurately. Do not claim that all calculations are sourced directly from one sheet unless the workbook has been checked.

## 🖱️ How to Use the Dashboard

1. Download or clone this repository to your computer.
2. Open the Excel workbook in Microsoft Excel.
3. If prompted, enable editing for this trusted local file.
4. Open the **Dashboard** worksheet.
5. Review the KPI cards to understand the overall data at a glance.
6. Use the available slicers or date filters to focus on a team, venue, toss decision, or other supported category.
7. Select or clear filters to compare different views.
8. Review the charts and match details to investigate the selected data.
9. Open **Tournament Highlights** for the tournament-wide summary.
10. If PivotTables are used and the source data changes, refresh the relevant PivotTables and confirm that the dashboard updates correctly.

> Slicers only control PivotTables, PivotCharts, and connections configured for them. Not every chart or KPI will respond to a slicer unless it is linked to the same data/filtering setup.

## 🧮 Methodology and Metric Definitions

The following are common definitions for this type of analysis. The actual workbook formulas should be reviewed to confirm which method is used.

### Winning Percentage

`Winning Percentage = (Matches Won / Matches Played) × 100`

Use a consistent match-counting method and account for ties, abandoned matches, or no-results if they exist in the dataset.

### Toss-to-Match Win Percentage

`Toss-to-Match Win % = (Matches Where Toss Winner Also Won / Eligible Matches) × 100`

Define the denominator consistently and exclude or handle no-results appropriately.

### Total Boundaries

If the source fields count fours and sixes separately:

`Total Boundaries = Fours + Sixes`

### Average Score

Average score should be calculated over the intended unit of analysis—for example, team innings or matches. Clearly identify which one is being displayed.

### Highest Team Score

The highest team score is the maximum valid team-innings score in the selected data.

### Distinct Counts

Count unique match IDs, teams, or venues where appropriate rather than relying on raw row counts if values can repeat.

## 🧹 Data Quality and Validation

Before interpreting results, check the source data for:

- Duplicate match records or duplicate team-innings records.
- Inconsistent team, city, or stadium names.
- Mixed or invalid date formats.
- Missing values in important fields.
- Numbers stored as text.
- Incorrect score, wicket, or margin values.
- Inconsistent categories for toss decisions and match results.

After cleaning or changing the source data, validate key totals and ensure that charts, formulas, PivotTables, and filters still behave as expected.

## 🗂️ Repository Structure

A suggested GitHub repository layout is:

```text
CWC-2023-Excel-Dashboard/
├── CWC2023.xlsx
├── README.md


Replace `CWC2023.xlsx` with the exact workbook filename in your repository. Add screenshots only if they exist in the repository.

## 🖼️ Screenshots

Add screenshots of the completed dashboard to make the project easier to review on GitHub.

For example, save a clear screenshot as `screenshots/dashboard.png`, then use this Markdown:

```markdown
![CWC 2023 Excel Dashboard](screenshots/dashboard.png)
```

You may also include screenshots of the toss analysis, venue analysis, and tournament highlights worksheets.

## ▶️ How to Run the Project

This is an Excel workbook project and does not require Python or a separate application to view.

1. Download `CWC2023.xlsx` from the repository.
2. Open the workbook in Microsoft Excel.
3. Navigate to the Dashboard worksheet.
4. Interact with the available slicers and filters.
5. Refresh PivotTables if the underlying source data has been updated.

Some features may behave differently in Excel for the web, LibreOffice, or other spreadsheet applications. Microsoft Excel desktop is recommended for the best compatibility with PivotTables, slicers, and workbook features.

## 💡 Findings and Interpretation

Use the dashboard to investigate questions such as:

- Which team recorded the most wins under the selected match scope?
- Which teams produced the highest scores?
- How did teams perform when batting first versus fielding first?
- How often did the toss winner also win the match?
- Which venues hosted the most matches?
- How did scoring patterns vary by venue?
- Which players delivered notable batting or bowling performances?

Record numerical findings here only after verifying them against the final workbook. The dashboard supports exploration of tournament data; observed relationships should not automatically be interpreted as causal relationships.

## ⚠️ Limitations

- Analysis quality depends on the completeness and accuracy of the source data.
- KPI values can differ depending on whether the calculation uses match-level, team-level, or innings-level records.
- Slicer behaviour depends on the connections configured in Excel.
- Tournament highlights are fixed historical facts and may not change with interactive filters.
- The dashboard describes the 2023 tournament and should not be treated as a prediction of future match outcomes.

## 🚀 Future Improvements

Possible extensions include:

- Add player-level batting and bowling analysis.
- Add more detailed venue comparisons.
- Improve accessibility and consistency of chart labels.
- Add a documented data dictionary and metric definitions.
- Automate data cleaning and refresh steps with Power Query.
- Add a dedicated validation page for checking totals and data quality.
- Extend the analysis to compare multiple tournaments if compatible data becomes available.

## 👤 Author

**Om Jadhav**  
Mechanical Engineering student and aspiring Data Analyst

This project is part of my learning portfolio and demonstrates my practice with Excel-based data analysis and dashboard design.


## 📜 Disclaimer

This project is intended for educational and portfolio purposes. Cricket statistics should be interpreted according to the source dataset and the definitions used in the workbook. ICC and Cricket World Cup names, logos, team names, and player names belong to their respective owners. This project is not an official ICC product and is not affiliated with or endorsed by the ICC.
