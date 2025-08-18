---
applyTo: '**'
---
Perfect — let’s draft this as a **Copilot instruction set** so it knows exactly what to do inside the swim tracker repo. I’ll frame it the same way we’ve been doing: clear goals, step-by-step implementation plan, formatting expectations, and commit guidelines.

---

# Attendance Weekly Summary Instructions

## 🎯 Goal

Add a new tab (`Attendance Summary`) to the MVHS Swim Tracker spreadsheet that shows **weekly attendance counts per swimmer**, flags whether each swimmer meets the **3 practices/week requirement**, and generates **charts split by sub-team** (4 charts total).

---

## 🛠 Implementation Plan

1. **Data Aggregation**

   * Create a new tab named `Attendance Summary`.
   * Aggregate attendance data from the `Attendance` tab by:

     * **Swimmer Name**
     * **Team / Sub-Team**
     * **Week Number** (use `ISOWEEKNUM(date)` or similar).
   * For each swimmer/week, count the number of “Present” (or equivalent) entries.

2. **Requirement Flag**

   * Add a column `Meets Requirement`:

     ```excel
     =IF([Attendance Count] >= 3, "✅", "❌")
     ```

3. **Pivot/Query Logic**

   * Use a `QUERY` or pivot table to generate the aggregated weekly data.
   * Ensure the query automatically updates when new attendance is added.
   * Example SQL-like query:

     ```sql
     select Swimmer, Team, ISOWEEKNUM(Date), count(Attendance)
     where Attendance = 'Present'
     group by Swimmer, Team, ISOWEEKNUM(Date)
     order by Team, Swimmer, ISOWEEKNUM(Date)
     ```

4. **Charting**

   * Insert 4 charts, one per sub-team.
   * Each chart:

     * X-axis = Week number
     * Y-axis = Attendance Count
     * Series = Each swimmer in that sub-team
   * Place charts neatly in the `Attendance Summary` tab, grouped by team.

5. **Team-Level Compliance Metrics**

   * Add a small table above each chart:

     * `% of swimmers meeting 3x/week`
     * `Average practices per swimmer per week`

---

## 🔧 Code & Formatting

* Follow existing Apps Script style conventions in this repo.
* Ensure all functions are added inside `CoachToolsCore` namespace if possible.
* Remove trailing spaces on new lines.
* Use spaces (not tabs) for indentation.
* Use 2-space indentation to match existing scripts.

---

## ✅ Commit Rules

* Stage and commit logical changes per feature:

  * One commit for data aggregation logic.
  * One commit for chart creation.
  * One commit for summary metrics (if added).
* Do not commit helper/test files or new markdown unless specifically requested.
* Commit message should reference the Jira story (if applicable) and clearly state:

  ```
  Added weekly attendance summary tab with per-swimmer counts, requirement flags, and team-level charts
  ```

---


