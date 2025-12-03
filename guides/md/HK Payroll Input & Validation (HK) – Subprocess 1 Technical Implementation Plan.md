# HK Payroll Input & Validation (HK) – Subprocess 1 Technical Implementation Plan



---

## 1. Scope & Objectives

**Subprocess 1 scope**

* Read all input workbooks from the Input folder (flexiform templates, One Time Payment, Inspire Awards, Employee Leave Transactions, EAO Summary, Workforce Detail, Merck Payroll Summary, SIP QIP, Flex Claim, RSU Dividend, AIP Payouts, 额外表, etc.).
* Create the master file `Flexi form out put YYYYMMDD.xlsx` with sheets:

  * `NewHire`, `InformationChange`, `SalaryChange`, `Termination`, `Attendance`, `VariablePay`.
* Fully copy raw data from flexiform templates into the first five sheets.
* Enrich `Attendance` and `VariablePay` by:

  * Joining and calculating leave days (Annual / Sick / Unpaid / PPTO / Maternity / Paternity).
  * Deriving EAO adjustment inputs.
  * Bringing in variable pay items (Lump Sum Merit, AIP, Inspire, SIP, RSU, Flex benefit, IA Pay Split, etc.).
  * Using config & 额外表 parameter tables (rates, policies, MPF/ORSO percentages, PPTO EAO Rate, etc.).
* Enforce **single-amount-per-employee-per-type** via `GroupByEmployeeAndType + SumAmount`.
* Ensure all **rounding and cross-month handling** is consistent and shared across the process.

**Technology constraints**

* **Control / Orchestration**: Power Automate Desktop (PAD).
* **Data processing**: Excel + VBA modules.
* **No Excel Power Query** (all ETL via VBA or plain formulas).
* All “heavy” Excel operations (loops, lookups, splits) should be done in **VBA**, not via PAD cell-level actions, for stability and performance.

---

## 2. Target Architecture (Subprocess 1)

### 2.1 Layered design

1. **PAD Orchestration Layer**

   * Triggered monthly (or on demand).
   * Reads runtime parameters (run date, payroll month, environment).
   * Opens the main **control workbook** and runs a single entry macro for Subprocess 1.
   * Handles high-level exception reporting (message box, log file, email).

2. **Excel “Control” Workbook (Macros)**

   * e.g. `HK_Payroll_Automation.xlsm`.
   * Contains:

     * All shared VBA services (dates, grouping, rounding, configuration, EAO, formatting).
     * Subprocess 1 orchestration macro: `Sub Run_Subprocess1()`.
   * Opens the various **input workbooks** and the **Flexi form out put** workbook.

3. **Input Workbooks**

   * All provided monthly source files (flexiform templates, One Time Payment, etc.).
   * Treated as **data only**; no business logic code inside.

4. **Output Workbook**

   * `Flexi form out put YYYYMMDD.xlsx`
   * Contains:

     * Raw copies of flexiform data.
     * Derived fields in `Attendance` & `VariablePay`.
   * Some helper named ranges to make formulas/macros robust.

### 2.2 Responsibilities: PAD vs VBA

| Concern                              | Owner | Notes                                                            |
| ------------------------------------ | ----- | ---------------------------------------------------------------- |
| Scheduling, trigger, environment     | PAD   | Windows Task Scheduler / PAD run                                 |
| Input folder / output folder paths   | PAD   | Passed via named cells or config                                 |
| Opening/closing Excel                | PAD   | One Excel instance per run                                       |
| Complex filters, joins, aggregations | VBA   | All business rules here                                          |
| Cross‑month splitting & holidays     | VBA   | Shared date service module                                       |
| MPF/ORSO/EAO calculations            | VBA   | Shared EAO & contribution services                               |
| Diff-summary formatting              | VBA   | Shared formatting service (called here or later by Subprocess 2) |

---

## 3. Configuration Strategy (No Hard‑Coding)

### 3.1 Configuration objects

All variable rules must come from **config workbooks**, not code:

* `config.xlsx` – core parameters:

  * `Calendar` sheet

    * Date, IsHKHoliday, IsBusinessDay (optional pre-computed).
  * `PayrollSchedule` sheet

    * Columns: `PayrollMonth` (YYYYMM), `CutoffDate`, `PayDate`, `IsAIPMonth`, `IsRSUDivMonth`, `IsRSUEYMonth`, `IsFlexBenefitMonth`, etc.
  * `ExchangeRates`

    * RateName (e.g. `RSU_Global`, `RSU_EY`, `DefaultFX`), RateValue.
  * `Final payment`, `Policy`, etc. (used more by Subprocess 2 but kept here to avoid later rework).
* `额外表.xlsx` – extra monthly parameters:

  * `[需要每月维护]` – e.g. PPTO EAO Rate input by WEIN.
  * `[MPF&ORSO]` – MPF / ORSO ratios.
  * `[特殊奖金]` – Flexible benefits, Other Allowance, etc.
  * `[Previous Month Terminated HC]`, `[Final payment]`, etc. (future use for Subprocess 2).

### 3.2 Config access service (VBA)

Create module `mConfig` with functions:

```vba
Type PayrollContext
    PayrollMonth As String   ' "YYYYMM"
    MonthStart As Date
    MonthEnd As Date
    PrevMonthStart As Date
    PrevMonthEnd As Date
    CutoffDate As Date
    PayDate As Date
End Type
```

Key functions:

* `GetPayrollContext(runDate As Date) As PayrollContext`

  * Determines target payroll month (e.g. via a named cell or mapping table).
  * Looks up `CutoffDate` and `PayDate` from `PayrollSchedule`.
* `GetExchangeRate(rateName As String) As Double`

  * Reads from `ExchangeRates`.
* `IsSpecialMonth(payrollContext, flagFieldName)`

  * Returns booleans like “Is RSU month?” so code never checks `Month(Date) = 5` directly.

**Usage examples**

* One time payment filter dates:

  * `CompletedOn` between `(PrevCutoffDate + 1)` and `CurrentCutoffDate`.
  * `ScheduledPaymentDate` in previous calendar month.
* RSU processing driven by:

  * `IsRSU_Global_Month` / `IsRSU_EY_Month` from config instead of month‑number checks.

---

## 4. Shared Service Library (VBA)

### 4.1 Rounding service (mandatory shared module)

Module: `mRounding`

Rules:

* **Monthly Salary**: round **once** to whole number and always use the rounded value in all subsequent calculations.
* **All calculation results** (pay items, adjustments, etc.): round to **two decimals**.

Functions:

```vba
Public Function RoundMonthlySalary(ByVal v As Variant) As Double
    RoundMonthlySalary = WorksheetFunction.Round(CDbl(v), 0)
End Function

Public Function RoundAmount2(ByVal v As Variant) As Double
    RoundAmount2 = WorksheetFunction.Round(CDbl(v), 2)
End Function

Public Function SafeAdd2(ByVal a As Variant, ByVal b As Variant) As Double
    SafeAdd2 = RoundAmount2(Nz(a) + Nz(b))
End Function
```

> All salary/amount fields must go through this module – **no direct `Round` scattered in business code**.

---

### 4.2 Grouping & Summation – `GroupByEmployeeAndType + SumAmount()`

Module: `mGrouping`

Goal: Reusable aggregator for all “if multiple records, sum first, then write one value per employee per type”.

Signature:

```vba
Public Function GroupByEmployeeAndType( _
    dataRange As Range, _
    employeeColName As String, _
    typeColName As String, _
    amountColName As String _
) As Object  ' returns Scripting.Dictionary
```

* Reads header row to resolve column indices (header-based, not index-based).
* Builds a key: `employeeID & "|" & typeValue`.
* Sums `amountCol` with `RoundAmount2` at the end.
* Returns: `Dictionary(key -> Double)`.

Convenience wrappers:

```vba
Public Function SumPerEmployee( _
    dataRange As Range, _
    employeeColName As String, _
    amountColName As String _
) As Object
```

Used in:

* One time payment: group by `(Employee ID, One-Time Payment Plan)`.
* Inspire Awards: `(Employee ID, One-Time Payment Plan)`.
* SIP QIP: `(EMPLOYEE ID, Pay Item)`.
* RSU: `(Employee Reference / EmployeeNumber, "RSU Global"/"RSU EY")`.
* AIP, Flex, etc.

This ensures the “only one amount per employee per plan type” rule is always respected and testable centrally.

---

### 4.3 Date & Calendar service (business days, HK public holidays)

Module: `mDateServices`

Key functions:

```vba
Public Function IsWeekend(d As Date) As Boolean
    IsWeekend = (Weekday(d, vbMonday) > 5)
End Function

Public Function IsHKPublicHoliday(d As Date) As Boolean
    ' Lookup in config.Calendar sheet
End Function

Public Function IsBusinessDay(d As Date) As Boolean
    IsBusinessDay = Not IsWeekend(d) And Not IsHKPublicHoliday(d)
End Function
```

#### Cross-month splitting primitives

We build generic helpers that all leave handlers reuse:

```vba
Public Type DateSpan
    StartDate As Date
    EndDate As Date
    YearMonth As String  ' "YYYYMM"
    Days As Double       ' semantics depend on leave type
End Type
```

Functions:

* `SplitByCalendarMonth(startDate, endDate) As Collection`

  * Splits an interval into multiple `DateSpan` entries (per calendar month).
* `CountBusinessDays(startDate, endDate) As Long`

  * Weeks + weekend/holiday exclusions.
* `SplitAnnualLeaveByMonthWithBusinessDays(...)`

  * Uses `IsBusinessDay` to compute **TOTAL DAYS excluding weekends + HK public holidays** for each month (per PDD).
* `HasFourConsecutiveBusinessDays(startDate, endDate) As Boolean`

  * For Sick Leave rule (>=4 consecutive working days).

These functions underpin the leave‑type specific modules.

---

### 4.4 Leave splitting & EAO services

Module: `mLeaveServices`

These functions rely on:

* `mDateServices` for splitting & business days.
* `mRounding` for numeric precision.
* `mConfig` for payroll context.

Conceptual responsibilities:

1. **Identify new unpaid leave records**
   Using composite key: `WIN|FROM_DATE|TO_DATE|APPLY_DATE|APPROVAL_DATE`.

   * Keep a simple **“history” sheet or file** storing keys already paid.
   * Each run:

     * Load history into a dictionary.
     * Select `STATUS="Approved"`.
     * Filter to records not in history → “unpaid” set.
     * After processing, append new keys to history.

2. **Leave type handlers**

Each type has a dedicated function returning a collection of **leave segments** ready to write to `Attendance` and/or `VariablePay`:

* `ProcessAnnualLeave(unpaidRecords, payrollContext, eaoSummaryIndex)`

  * Splits by month **excluding weekends & HK holidays**.
  * Segments:

    * Current month → `Attendance.Days_AnnualLeave` & `.Days_AnnualLeaveForDeduction`.
    * Previous month → `Attendance.Days_AnnualLeave_LastMonth` & `_ForDeduction_LastMonth`.
    * Before previous month → aggregate days → EAO: `(AverageDayWage_12Month - DailySalary) * TOTAL_DAYS` and write to `VariablePay.Annual Leave EAO Adj_Input`.

* `ProcessSickLeave(...)`

  * Discards ranges without **4 consecutive business days**.
  * Splits by month (calendar days).
  * Current & previous month days go into Attendance (current/last month & ForDeduction columns).
  * Older periods trigger EAO: `(DayWage_Maternity/Paternity/Sick Leave − DailySalary) * Days_SickLeave`, aggregated per WEIN and mapped to `VariablePay.Sick Leave EAO Adj_Input`.

* `ProcessUnpaidLeave(...)`

  * Splits by month (calendar).
  * Current + previous month days to `Attendance` (No Pay Leave & LastMonth).
  * Older → EAO: `NoPayLeaveCalculationBase * Days_NoPayLeave`, aggregated per WEIN; added into `VariablePay.No Pay Leave Deduction` (if cell non‑empty, **add**, not overwrite).

* `ProcessPPTO(...)`

  * Splits by month (calendar).
  * Applies “current month payroll settles previous month” logic via `PayrollContext`.
  * Current month portion → `Attendance.Days_Paid Parental Time Off` (+ ForDeduction).
  * Previous month portion → `_LastMonth` columns.
  * Future segments kept in **carryover file/sheet** for next months.

* `ProcessMaternityLeave(...)` / `ProcessPaternityLeave(...)`

  * Additional check on **40 weeks service** for Maternity (via Workforce Detail).
  * If <40 weeks:

    * Exclude from payroll settlement.
    * Write to separate `Maternity report` workbook.
  * Else:

    * Apply same “current month settles previous month” logic as PPTO.
    * Older segments → EAO `(DayWage_Maternity/Paternity/Sick Leave − DailySalary) * Days` with grouping per WEIN; mapped to VariablePay.

All EAO formulas and ratio fields come from EAO Summary & config tables, not code.

---

### 4.5 Diff-column summary formatting service

Even though **Diff columns** are output in Subprocess 2, the logic fits naturally as a reusable VBA service.

Module: `mFormatting`

```vba
Public Sub SummarizeDiffColumns(diffHeaderRow As Long, _
                                firstDiffCol As Long, _
                                lastDiffCol As Long, _
                                firstDataRow As Long, _
                                lastDataRow As Long)
    Dim col As Long, falseCount As Long
    For col = firstDiffCol To lastDiffCol
        falseCount = WorksheetFunction.CountIf( _
                         Range(Cells(firstDataRow, col), Cells(lastDataRow, col)), _
                         "FALSE")
        Cells(diffHeaderRow, col).Value = falseCount
        If falseCount > 0 Then
            Cells(diffHeaderRow, col).Interior.Color = vbRed
        Else
            Cells(diffHeaderRow, col).Interior.ColorIndex = xlColorIndexNone
        End If
    Next col
End Sub
```

* Called from Subprocess 2’s validation macro.
* Kept in shared module so Subprocess 1 (or any other workbook) can reuse if needed.

---

## 5. PAD Flow Design – Subprocess 1

### 5.1 PAD input parameters

PAD variables:

* `vRunDate` – “today” according to scheduler.
* `vInputFolder`, `vOutputFolder`, `vConfigFolder`.
* `vLogFolder`.

PAD writes `vRunDate` and `targetPayrollMonth` into named cells in `config.xlsx` (or control workbook), e.g.:

* `Config!B2 = RunDate`
* `Config!B3 = PayrollMonth` (YYYYMM string).

### 5.2 High-level PAD steps

1. **Initialization**

   * Load configuration paths from PAD settings / environment variables.
   * Create a new run log file (`YYYYMMDD_HHMM_Subprocess1.log`).

2. **Input file discovery & checks**

   * For each expected source pattern (e.g. `1263 ADP flexiform template_HK_NewHire*.xlsx`, `One time payment report*.xlsx`, etc.):

     * Use PAD “Get files in folder” + filter by pattern.
     * If missing or >1 result, log error and stop (or mark as warning if business confirms optional).

3. **Launch Excel and control workbook**

   * Start Excel instance.
   * Open `HK_Payroll_Automation.xlsm`.
   * Ensure macros are enabled.

4. **Run Subprocess 1 macro**

   * Use PAD action “Run Excel macro” with macro name:

     * `Run_Subprocess1`
   * Pass necessary paths via:

     * Named ranges in control workbook (e.g. `NamedRange_InputFolder`, `NamedRange_OutputFolder`).
     * Or by writing them into a `RuntimeParameters` sheet.

5. **Error handling**

   * Macro returns success/failure flag via a named cell or a tiny result sheet.
   * PAD reads the result; if failure:

     * Include error summary from log sheet in email to SME.

6. **Save & close**

   * Ensure `Flexi form out put YYYYMMDD.xlsx` is saved in `OutputFolder`.
   * Close all workbooks and Excel.

7. **Notification**

   * PAD sends a simple notification (Outlook, Teams, etc.) with run status and path to output file.

---

## 6. VBA Implementation – Subprocess 1 Workflows

### 6.1 Main entry macro

In `HK_Payroll_Automation.xlsm`:

```vba
Public Sub Run_Subprocess1()
    On Error GoTo ErrHandler

    Dim ctx As PayrollContext
    Dim flexWb As Workbook

    ' 1. Get runtime context and config
    Set ctx = GetPayrollContext(Range("Config_RunDate").Value)

    ' 2. Create output workbook
    Set flexWb = CreateFlexiOutputWorkbook(ctx)

    ' 3. Copy raw flexiform data into relevant sheets
    LoadFlexiformData flexWb, ctx

    ' 4. Populate Attendance sheet (leave days & adjustments)
    PopulateAttendance flexWb, ctx

    ' 5. Populate VariablePay sheet (variable pay, EAO inputs, etc.)
    PopulateVariablePay flexWb, ctx

    ' 6. Final formatting & save
    FinalizeFlexOutput flexWb, ctx

    Exit Sub

ErrHandler:
    ' Log error to dedicated sheet and bubble status to PAD
End Sub
```

---

### 6.2 Create `Flexi form out put YYYYMMDD.xlsx`

`CreateFlexiOutputWorkbook(ctx)`

* Uses `ctx.PayDate` or `ctx.MonthEnd` to build filename `Flexi form out put YYYYMMDD.xlsx`.
* Creates workbook with sheets:

  * `NewHire`, `InformationChange`, `SalaryChange`, `Termination`, `Attendance`, `VariablePay`.
* Sets up:

  * Table headers (copied from flexiform templates where necessary).
  * Named ranges for major blocks (e.g. `tblAttendance`, `tblVariablePay`).
* **No business data written here**; just structure.

---

### 6.3 Copy raw flexiform data

`LoadFlexiformData flexWb, ctx`

For each flexiform file:

* Use header-based copy:

  * Open `1263 ADP flexiform template_HK_NewHire` etc.
  * Identify the main data table range.
  * Copy entire header + data to `flexWb.Sheets("NewHire")`.

Repeat for:

* `DataChange` → `InformationChange`.
* `Comp` → `SalaryChange`.
* `Termination` → `Termination`.
* `Attendance` → `Attendance`.
* `Variable` → `VariablePay` base data.

Then:

* On `VariablePay`, insert six new columns to the right of **“Adjustment of Parental Paid Time Off (PPTO) payment”**:

  * `IA Pay Split`
  * `MPF Relevant Income Rewrite`
  * `MPF VC Relevant Income Rewrite`
  * `Paternity Leave payment adjustment`
  * `PPTO EAO Rate input`
  * `Flexible benefits`

Headers are constant and used later by business logic.

---

### 6.4 Populate `Attendance` sheet (shared leave services)

`PopulateAttendance flexWb, ctx`

Steps:

1. **Load Employee Leave Transactions**

   * Open `Employee_Leave_Transactions_Report`.
   * Read into in-memory array / recordset (performance).
   * Filter `[STATUS = Approved]`.

2. **Identify new unpaid records**

   * Build unique key `WIN|FROM_DATE|TO_DATE|APPLY_DATE|APPROVAL_DATE`.
   * Compare vs history (sheet in control workbook or dedicated history workbook).
   * Keep only unpaid records.

3. **Dispatch by LEAVE TYPE**

   * Group unpaid records by leave type.
   * Call per-type handlers in `mLeaveServices`:

     * `ProcessAnnualLeave` → returns:

       * Per WEIN: days for current month, previous month, older months.
     * `ProcessSickLeave`
     * `ProcessUnpaidLeave`
     * `ProcessPPTO`
     * `ProcessMaternityLeave`
     * `ProcessPaternityLeave`

4. **Write back to `Attendance`**

   * For each WEIN (mapped by Employee Code):

     * Find or append row in `Attendance`:

       * If row exists, update columns.
       * If it doesn’t exist yet, append a new row (Employee Code + data per PDD rule).
   * Use header-based column detection to avoid brittle column indices.

5. **Update history**

   * Append processed keys to “leave history” sheet.

All date logic uses `mDateServices` and `PayrollContext` so changing payroll schedule (cut-off or pay date) doesn’t require code changes.

---

### 6.5 Populate `VariablePay` sheet

`PopulateVariablePay flexWb, ctx`

High‑level pattern:

1. **Create in-memory index of `VariablePay` by Employee Code**

   * Dictionary: `employeeCode → rowNumber`.
   * For new employees sourced only from other input tables, append row and update dictionary.

2. **Call data‑source specific loaders**

For each data source, we:

* Load to memory.
* Apply config-driven filters (dates, months, statuses).
* Use `GroupByEmployeeAndType` where applicable.
* Map aggregated results into target columns of `VariablePay`.

Examples:

#### 6.5.1 One time payment report

* Filter out `[One-Time Payment Plan] = Inspire Points Value / Inspire Cash Award`.
* Filter `[Completed On]` using `ctx`:

  * `CompletedOn > PrevCutoffDate` and `<= CurrentCutoffDate`.
* Filter `[Scheduled Payment Date]` ∈ previous calendar month (using date service).
* Use `GroupByEmployeeAndType` on `(Employee ID, One-Time Payment Plan)` → `Actual Payment – Amount`.

Map plan names to columns:

| One-Time Payment Plan              | VariablePay column          |
| ---------------------------------- | --------------------------- |
| Lump Sum Merit                     | `Lump Sum Bonus`            |
| Sign On Bonus                      | `Sign On Bonus`             |
| Retention Bonus                    | `Retention Bonus`           |
| Referral Payment                   | `Referral Bonus`            |
| Manager of the Year Award          | `Manager of the Year Award` |
| MD Award                           | `MD Award`                  |
| Employee Award                     | `Employee Award`            |
| New Year's Allowance / Red Packet  | `Red Packet`                |
| Cash award / SIP to AIP Transition | `Other Allowance`           |

* All amounts: `RoundAmount2`.
* Per employee-plan, only a single final value is written.

#### 6.5.2 Inspire Awards payroll report

* Filter `One-Time Payment Plan = Inspire Points Value` and `Inspire Cash`.
* Group by `(Employee ID, Plan)` → `Actual Payment – Amount`.
* Map to columns `InspirePoints` and `Inspire Cash` on `VariablePay`.

#### 6.5.3 EAO-based adjustments from older leave

Use results from `mLeaveServices`:

* Annual Leave EAO Adj_Input.
* Sick Leave EAO Adj_Input (sum per WEIN).
* No Pay Leave Deduction (sum and **add** to existing cell if non-zero).
* Maternity Leave EAO Adj_Input.
* Paternity Leave EAO Adj_Input.

All formulas & rates taken from `EAO Summary Report_YYYYMM` and config.

#### 6.5.4 Merck Payroll Summary Report——xxx

* For each employee’s report:

  * Extract `Employee ID`, `Net Pay (include EAO & leave payment)`, `MPF Relevant Income`, `MPF VC Relevant Income`, `MPF EE MC`, `MPF EE VC`.
* Compute:
  `IA Pay Split = Net Pay (include EAO & leave payment) + MPF EE MC + MPF EE VC`.
* Map:

  * `IA Pay Split` → `VariablePay.IA Pay Split`.
  * `MPF Relevant Income` → `VariablePay.MPF Relevant Income Rewrite`.
  * `MPF VC Relevant Income` → `VariablePay.MPF VC Relevant Income Rewrite`.
* Use `RoundAmount2`.

#### 6.5.5 SIP QIP

* Filter `Pay Item = Qualitative Incentive Plan` and `Sales Incentive Plan`.
* Group by `(EMPLOYEE ID, Pay Item)` → `TOTAL PAYOUT`.
* Map to:

  * `Sales Incentive (Qualitative)` and
  * `Sales Incentive (Quantitative)`.

#### 6.5.6 MSD HK Flex_Claim_Summary_Report

* Check config: `IsFlexBenefitMonth(ctx)` (instead of checking Feb/Aug directly).
* If true:

  * Filter `Claim Status = Approved`.
  * Group by `Employee Number ID` → `Transacted Amount`.
  * Map to `VariablePay.Flexible benefits`.

#### 6.5.7 RSU Dividend

* For RSU Global:

  * Check `IsRSU_Global_Month(ctx)` – normally May.
  * If true:

    * Extract `(Employee Reference, Gross Award Amount to be Paid)`.
    * Get FX from `GetExchangeRate("RSU_Global")`.
    * value = `Gross Award × FX`.
    * Group by Employee Reference.
    * Map to `VariablePay.Shares Dividend`.

* For RSU EY:

  * Check `IsRSU_EY_Month(ctx)` – normally June.
  * Similar pattern with `Dividend To Pay × FX`.

#### 6.5.8 AIP Payouts Payroll Report

* Check `IsAIPMonth(ctx)` – normally March.
* If true:

  * Group by Employee WIN → `Bonus Amount`.
  * Map to `VariablePay.Annual Incentive`.

#### 6.5.9 PPTO EAO Rate input & Flexible benefits from 额外表

* From `额外表.[需要每月维护]`:

  * Map `(WEIN → PPTO EAO Rate input)` to `VariablePay.PPTO EAO Rate input`.

* From `额外表.[特殊奖金]`:

  * `Flexible benefits`, potentially additional allowances (if required by Subprocess 1).

All mapping uses the **same Employee Code index** created at the beginning.

---

### 6.6 Finalization & sanity checks

`FinalizeFlexOutput flexWb, ctx`

* Apply standard formatting:

  * Freeze panes on header rows.
  * Autofit columns.
  * Number formats:

    * Monetary fields: `#,##0.00`.
    * Days: `0.00` or `0.0` per SME preference.
  * Date fields: `yyyy/mm/dd`.

* Optionally:

  * Insert a small **run summary** sheet:

    * PayrollMonth, CutoffDate, PayDate.
    * Number of records in each sheet.
    * Number of employees in Attendance / VariablePay.

* Save in Output folder with final filename.

---

## 7. Error Handling, Logging, and Idempotency

* **Central logging** module writes messages to:

  * `Log` sheet in control workbook, and/or
  * A text log file per run.
* Each major step (per data source) logs:

  * Start time, records read, records written, error counts.
* Any row-level exception is logged with:

  * Employee ID, source file, sheet, error message.
* **Idempotency**:

  * Key concept for leave processing and EAO:

    * Using the persistent **leave history** ensures re-runs do not double-count older leave records.
  * For variable pay (e.g. One time payment, Inspire, RSU):

    * The output always completely recalculates from inputs rather than incrementally appending.

---

## 8. Testing Strategy (Subprocess 1)

* **Unit tests for shared services**:

  * `GroupByEmployeeAndType` on synthetic tables to confirm grouping per employee & type.
  * `IsBusinessDay` and `IsHKPublicHoliday` with known calendar.
  * `SplitAnnualLeaveByMonthWithBusinessDays` against worked examples (e.g. 09/30–11/01 case).
  * Sick Leave: test 3 vs 4 consecutive working day scenarios.
  * Rounding: confirm monthly salary rounding and 2-decimal rounding.

* **Integration tests per data source**:

  * Build small representative input files:

    * Multiple entries for same employee & plan.
    * Cross-month leave ranges (including PHs and weekends).
    * Leave before previous month to trigger EAO adjustments.
  * Verify:

    * `Attendance` days distribution per month.
    * `VariablePay` mapping per plan type and per employee.

* **Regression baseline**:

  * For one full real payroll month, compare the automated `Flexi form out put` against a manually validated version, and store as a regression benchmark.

---

This plan keeps **Subprocess 1** highly modular, PAD-light and VBA-heavy where appropriate, uses configuration for all business rules (dates, rates, policy switches), and centralizes reusable services (grouping, dates, EAO, rounding, and Diff-formatting) so Subprocess 2 can be layered on later without redesigning the foundations.
