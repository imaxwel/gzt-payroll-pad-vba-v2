# HK Payroll Input & Validation (HK) – Subprocess 2

**Technical Implementation Plan (PAD + Excel VBA)**



---

## 1. Scope & Objectives

**Subprocess 2 scope (“VALIDATION part”)**

Per the PDD, Subprocess 2: 

* Reads multiple input tables (current & previous Payroll Report, Workforce Detail – Payroll‑AP, flexiform NewHire & Termination, Employee_Leave_Transactions_Report, EAO Summary Report_YYYYMM, Allowance plan report, 2025QX Payout Summary, Inspire Awards payroll report, RSU Dividend reports, One time payment report, Merck Payroll Summary Report——xxx, Optional medical plan enrollment form, 额外表, etc.).
* Creates **`HK Payroll Validation Output YYYYMMDD.xlsx`** with:

  * **`Check Result`** sheet:

    * Benchmark section = full copy of **Payroll Report**.
    * Check columns = values recalculated / re‑sourced from all input tables.
    * Diff columns = TRUE/FALSE comparison between Benchmark & Check.
    * First row summary = FALSE counts per Diff column, red if >0.
  * **`HC Check`** sheet:

    * Headcount cross‑check by Hire Status, terminations, new hires, 额外表, etc.
* Applies business rules for:

  * Master data (names, dates, org info, salary, transport allowance).
  * Pay items (Base Pay, leave payments, AIP, sales incentives, Inspire, RSU, Red Packet, year‑end, lump sums, etc.).
  * Final payment (Severance, Long Service, PIL, Gratuities, Back Pay).
  * Contributions (MPF relevant income / VC / EE/ER MC & VC, ORSO, etc.).
  * Benefits for tax (Inspire Points gross‑up).

**Technology & constraints**

* Orchestration: **Power Automate Desktop (PAD)**.
* Data processing & business rules: **Excel VBA** (in macro‑enabled workbook).
* **No Power Query** – all ETL and comparison logic must be VBA / native Excel only.
* Shared services from Subprocess 1 (calendar, grouping, rounding, config, EAO, formatting) must be reused to avoid drift.

---

## 2. Overall Architecture

### 2.1 PAD orchestration

Single PAD flow for Subprocess 2, e.g. `HK_Payroll_SP2_Main`:

1. **Load runtime parameters**

   * Input folder, output folder, config folder (from environment or `config.xlsx`).
   * Run date (today); target **payroll month** (either from user input or config).

2. **Input file discovery**

   * Required files (logical names from config):

     * `PayrollReport_Current`, `PayrollReport_Previous`
     * `WorkforceDetail`, `Flex_NewHire_Previous`, `Flex_NewHire_Current`, `Flex_Termination_Previous`, `Flex_Termination_Current`
     * `EmployeeLeaveTransactions`, `EAO_Summary_Current`, `EAO_Summary_Previous`
     * `AllowancePlan`, `2025QX_Payout_Summary`, `InspireAwards`, `RSU_Global`, `RSU_EY`, `OneTimePayment`, `MerckPayrollSummary_*`, `OptionalMedical`, `额外表`, `HK_Payroll_Validation_Output_Template` etc. 
   * Validate **mandatory** vs **optional** sources (from config table).

3. **Launch Excel**

   * Start one Excel instance.
   * Open **macro control workbook** (e.g. `Payroll_HK_Macros.xlsm`).
   * Open `config.xlsx` and `额外表.xlsx` if not already loaded by macros.

4. **Run Subprocess 2 macro**

   * Call `SP2_Run` entry point, passing:

     * `InputFolderPath`
     * `OutputFolderPath`
     * `RunDate`
     * `PayrollMonth`
   * PAD waits until macro finishes.

5. **Result handling**

   * Macro writes status + message into a known cell / sheet.
   * PAD reads it; if failed, captures error log and notifies team.
   * If success, keeps `HK Payroll Validation Output YYYYMMDD.xlsx` in output folder and sends notification.

6. **Cleanup**

   * Close all workbooks and Excel.
   * Archive input if needed.

**Rule of thumb:** PAD handles **file logistics & scheduling**; all domain logic lives in VBA.

---

### 2.2 VBA solution structure

Use the same macro workbook as Subprocess 1 (or a closely related one), with shared modules:

* **Shared from Sub1**

  * `modConfigService` – config & parameter reading (config.xlsx + 额外表).
  * `modCalendarService` – payroll periods, business days, HK holidays.
  * `modGroupingService` – `GroupByEmployeeAndType + SumAmount`.
  * `modRoundingService` – monthly salary rounding & 2‑decimal rule.
  * `modLookupService` – WEIN/Employee ID/Employee Code mapping & append logic.
  * `modEAOService` – EAO calculations.
  * `modFormattingService` – Diff summary formatting.

* **Sub2‑specific**

  * `modSP2_Entry` – `Sub SP2_Run(...)`.
  * `modValidationOutput` – create `HK Payroll Validation Output` file, layout Check/ Diff columns.
  * `modCheckResult_MasterData` – Legal names, hire/term dates, org info, base pay, transport allowance.
  * `modCheckResult_PayItems` – Base Pay, Salary Adj, Transport Allowance Adj, leave payments, AIP, etc.
  * `modCheckResult_Incentives` – SIP/2025QX, Inspire, RSU, lump sums, Red Packet, Flexible benefits.
  * `modCheckResult_FinalPayment` – Severance, Long Service, PIL, Gratuities, Back Pay, year‑end bonus.
  * `modCheckResult_Contribution` – MPF, ORSO, related base calculations.
  * `modCheckResult_BenefitsTax` – Inspire Points gross‑up.
  * `modDiffEngine` – copy benchmark → Check; compute Diff TRUE/FALSE; special rules (Last Hired Date pre‑2025).
  * `modHCCheck` – pivot, headcount checks with previous/current month & 额外表.

Each module should be **header‑driven**, not dependent on hard‑coded column indices.

---

## 3. Configuration & Parameterisation

We extend the Sub1 config strategy to Sub2.

### 3.1 `config.xlsx` additions

* **GlobalConfig**

  * Add:

    * `ValidationTemplateFile` → path/filename of `HK payroll validation output template`.
    * `HC_Pivot_RowField` (e.g. `Hire Status`).
    * `HC_Pivot_ValueField` (e.g. `WEIN`).

* **PayrollCalendar**

  * Already used for cut‑off and pay dates; reused for:

    * Determining **Current Month** vs **Previous Month** payroll reports.
    * Headcount rules referencing *“Termination Date + 7 vs pay day”*. 

* **FinalPayment** (either in config or 额外表, per PDD)

  * Columns:

    * `WEIN`, `MSD_or_Statutory`, `TerminationType`, `PILIndicator`, `NoticeGivenDate`, `NoticePeriod`, `Gratuities`, `BackPay`, etc. 

* **RoundingPolicy**

  * Document that:

    * **Monthly Salary** is rounded **to integer** immediately after extraction.
    * **All calculation results** default to **2 decimals**, except where PDD requires integer (e.g. certain headcount & gross‑up values).

### 3.2 `额外表.xlsx` for Sub2

Already defined in Sub1 but used more intensely here:

* `[MPF&ORSO]` – MPF EE/ER VC %, ORSO %, ORSO ER Adj, Percent of ORSO EE/ER, etc. 
* `[Final payment]` – policy type MSD vs Statutory, parameterised formulas, PIL indicator & notice periods.
* `[Previous Month Terminated HC]` – lists WEIN for previous / current month to feed HC Check.
* `[特殊奖金]` – Flexible benefits, Other Allowance, Other Bonus, Other Rewards, Goods & Services Differential, etc.

All are treated as **parameter tables**; macros never hard‑code thresholds or percentages.

### 3.3 HK payroll validation output template

* Static template `HK Payroll Validation Output.xlsx` defines: 

  * Check Result header structure (Benchmark / Check / Diff columns & order).
  * HC Check layout (where to place pivot & summary values).

`modValidationOutput` will:

* Open this template.
* Copy header rows & structural formatting into the new `HK Payroll Validation Output YYYYMMDD.xlsx`.
* Keep a **metadata table** (in hidden sheet or in code) linking:

  * For each logical field (e.g. `Monthly Base Pay`):

    * Benchmark column name.
    * Check column name.
    * Diff column name.

This metadata powers the generic Diff engine (no hard‑coded column letters).

---

## 4. PAD Flow for Subprocess 2

Recommended steps:

1. **Initialisation**

   * Set PAD variables for run date, input/output/config folders.
   * (Optionally) read `PayrollMonth` and `PreviousPayrollMonth` from config or from user input.

2. **File discovery & validation**

   * Use `Get files in folder` with patterns (from `InputFileMapping` sheet in `config.xlsx`) to locate:

     * Current & previous Payroll Report (`Payroll Report*.xlsx`).
     * Current Workforce Detail.
     * Current & previous NewHire and Termination flexiforms.
     * EAO Summary (current and relevant old months if needed).
     * Allowance plan report, 2025QX Payout Summary, One time payment report, etc. 
   * If any **mandatory** source missing, log and stop with clear error status.

3. **Launch Excel / run macros**

   * Open macro workbook.
   * Write resolved paths into a small runtime sheet or named ranges.
   * Run `SP2_Run`.

4. **Post‑run**

   * Read status cell.
   * If success:

     * Confirm that `HK Payroll Validation Output YYYYMMDD.xlsx` exists in output folder.
     * Send notification including path.
   * If failure:

     * Attach log (e.g. log sheet exported to .csv) and notify RPA support.

No PAD cell‑by‑cell manipulation; everything inside Excel is macro‑driven.

---

## 5. VBA Implementation Plan for Subprocess 2

### 5.1 Entry point

```vb
Public Sub SP2_Run( _
    ByVal inputFolder As String, _
    ByVal outputFolder As String, _
    ByVal runDate As Date, _
    ByVal payrollMonth As String)

    On Error GoTo ErrHandler

    Dim ctx As tPayrollCalendar
    ctx = GetPayrollContext(runDate)     ' from modCalendarService

    ' 1. Create HK Payroll Validation Output workbook
    Dim valWb As Workbook
    Set valWb = CreateValidationOutputWorkbook(outputFolder, runDate)

    ' 2. Build Check Result benchmark + WEIN index
    BuildBenchmarkAndIndex valWb, inputFolder, payrollMonth, ctx

    ' 3. Populate Check columns (modular groups)
    RunMasterDataChecks valWb, inputFolder, ctx
    RunPayItemChecks valWb, inputFolder, ctx
    RunIncentiveChecks valWb, inputFolder, ctx
    RunFinalPaymentChecks valWb, inputFolder, ctx
    RunContributionChecks valWb, inputFolder, ctx
    RunBenefitsForTaxChecks valWb, inputFolder, ctx

    ' 4. Run Diff engine
    ComputeDiffs valWb, ctx

    ' 5. Build HC Check sheet
    BuildHCCheck valWb, inputFolder, ctx

    ' 6. Save & status
    SaveValidationOutput valWb, outputFolder, runDate

    ' Write success status for PAD
    ThisWorkbook.Worksheets("Runtime").Range("SP2_Status").Value = "OK"
    Exit Sub

ErrHandler:
    LogError "SP2_Run", Err.Number, Err.Description
    ThisWorkbook.Worksheets("Runtime").Range("SP2_Status").Value = "ERROR"
End Sub
```

This is **planning‑level pseudocode**, not final code.

---

### 5.2 Create `HK Payroll Validation Output YYYYMMDD.xlsx`

`CreateValidationOutputWorkbook`:

1. Open template file (from `GlobalConfig.ValidationTemplateFile`).
2. Save‑As new workbook named `HK Payroll Validation Output YYYYMMDD.xlsx` in output folder. 
3. Ensure it contains:

   * `Check Result` sheet with HK template header (Benchmark + Check + Diff).
   * `HC Check` sheet layout.

No data is populated yet.

---

### 5.3 Build Benchmark & WEIN index

`BuildBenchmarkAndIndex valWb, inputFolder, ctx`

**Inputs**: current month `Payroll Report`.

According to PDD: copy **raw benchmark** from Payroll Report into Check Result, preserving all columns (and some may be dynamic). 

Steps:

1. Open current `Payroll Report` workbook.
2. Identify **data table range** (header row + rows with data).
3. Copy entire table to `Check Result` starting at fixed anchor (e.g. `A4`) – this is the **Benchmark Data** section.
4. Build **`WEINIndex` dictionary**:

   * Key = WEIN from Payroll Report.
   * Value = row number in `Check Result` for that WEIN.
5. Store `WEINIndex` in a module‑level object so all check modules can re‑use it.

Later, when other sources contain WEIN that is **not in benchmark**, they will:

* Append new row to **end** of Benchmark Data:

  * Copy header structure (blank for Benchmark columns).
  * Set WEIN cell and any required minimal fields.
* Update `WEINIndex` for the new row.

This implements PDD rule:

> If data exists in other tables but not in the benchmark table, append it to the end of Benchmark Data and fill Check columns accordingly, leaving Benchmark blank. 

---

## 6. Check Result – Check Columns Implementation

Implementation follows the order in the PDD, but grouped in logical blocks.

### 6.1 Common pattern for all Check modules

For each block (Master Data, Pay Items, etc.):

1. **Pre‑load source(s)** into in‑memory arrays / dictionaries:

   * Workforce Detail → keyed by Employee ID.
   * Payroll Report prev month → keyed by WEIN.
   * 额外表 sheets → keyed by WEIN.
   * etc.
2. **Bridge WEIN ⇔ Employee ID ⇔ Employee Code** using `modLookupService`.
3. For each relevant WEIN in `WEINIndex`:

   * Compute / retrieve value.
   * Write into **Check** column in `Check Result`.
4. For WEINs present in source but not in `WEINIndex`:

   * Use `MapOrAppendByWEIN` helper:

     * Appends new row at bottom.
     * Updates index & writes Check value there.

This pattern avoids inconsistent behaviour across blocks.

---

### 6.2 HC Check sheet (headcount logic)

Although described under Sub2 in PDD, HC Check is logically separate; implement in `modHCCheck`. 

**Sources**:

* Current & previous month `Payroll Report`.
* Current & previous month `flexiform_HK_Termination`.
* Current & previous month `flexiform_HK_NewHire`.
* `额外表.[Previous Month Terminated HC]`.

**Steps**:

1. **Pivot headcount on HC Check**

   * Use current month Payroll Report as data source.
   * Create PivotTable:

     * Rows: Hire Status.
     * Values: Count of WEIN.
   * Place on `HC Check` at template‑specified location.

2. **Previous & current active HC**

   * From previous month Payroll Report, filter `[Hire Status=Active]` and count WEIN; write value to the cell indicated as “Last Month Payroll HC”.
   * Same for current month; write to “Current Month Payroll HC”.

3. **Terminated HC classification**

   * From previous month flex Termination:

     * For each Employee Code, get Termination Date.
     * Determine pay day from `ctx` (business day‑adjusted).
     * If `TerminationDate + 7 > PayDay` → classify as **Current Month Terminated HC (included)**.
     * Else → **Current Month Terminated HC (OC)**.
     * Write counts to previous month row.
   * Repeat for current month flex Termination (writing to current month row). 

4. **Previous Month Terminated HC from 额外表**

   * From `[Previous Month Terminated HC]` sheet:

     * Count WEIN for previous and current month.
     * Write to designated red/green cells.

5. **New hire counts**

   * Count Employee ID in previous & current NewHire flex for each month; write to red/green cells.

6. **Check column formula**

   * In adjacent “Check” column, insert formula:

     * `LastMonthPayrollHC - LastMonthTerminatedIncluded - CurrentMonthTerminatedOC + CurrentMonthNewHC`
   * Compare result with current Payroll active HC; optionally add a simple Diff column for HC.

All cell addresses should be driven by named ranges / template metadata, not hard‑coded coordinates.

---

### 6.3 Master Data Check

Per PDD (Legal Full Name, Last Hired Date, Last Employment Date, Business Department, Position Title, Cost Center, Monthly Base Pay, Monthly Transport Allowance). 

Module `modCheckResult_MasterData`:

**Pre‑load:**

* `Workforce Detail - Payroll-AP`:

  * Employee ID, WEIN, Legal First/Last Name, Last Hire Date, Business Department, Position Title, Cost Center - ID, Monthly Salary, Employee Type.
* `flex_Termination` (current & previous, as needed): Employee ID, Termination Date.
* `Allowance plan report`: Employee ID, Amount where Compensation Plan = Transportation Allowance.

**Calculations:**

1. **Legal Full Name**

   * Benchmark column: concatenation `Legal First Name & " " & Legal Last Name` (from Payroll Report).
   * Check column:

     * For each WEIN, map to Employee ID, then look up **Legal Full Name** from Workforce Detail.
     * Write into `Legal full name Check`.
2. **Last Hire Date Check**

   * Workforce Detail → Last Hire Date; map to WEIN; write.
   * **Special Diff rule**: if either Benchmark or Check Last Hired Date is **before 2025‑01‑01**, Diff should be **TRUE** regardless (applied by Diff engine).
3. **Last Employment Date Check**

   * Collect Termination Date from flex Termination; map Employee ID → WEIN.
4. **Business Department / Position Title / Cost Center Code Check**

   * All from Workforce Detail via Employee ID → WEIN mapping.
5. **Monthly Base Pay / Monthly Base Pay (Temp) Check**

   * From Workforce Detail:

     * Filter `[Employee Type=Regular]` → Monthly Base Pay.
     * `[Employee Type=Intern/Co-ops]` → Monthly Base Pay (Temp).
   * Immediately apply `RoundMonthlySalary` from `modRoundingService`.
6. **Monthly Transport Allowance Check**

   * From Allowance plan report:

     * `[Compensation Plan = Transportation Allowance]` → Amount.
   * Write to `Monthly Transport Allowance Check`.

Rounding: monthly salary **integer**; other numeric Check fields use 2‑decimal rule where required.

---

### 6.4 Pay Items & Leave‑related Checks

According to PDD, Pay Items block includes: Base Pay / Base Pay (Temp), Salary Adj / Transport Allowance Adj, Transport Allowance, Total EAO Adj, various leave payments, No Pay Leave Deduction, Untaken Annual Leave Payment, etc. 

Module: `modCheckResult_PayItems`.

**Pre‑load:**

* Workforce Detail – Monthly Salary (rounded), Last Hire Date.
* Employee_Leave_Transactions_Report (approved only, with composite key for new/unpaid events).
* EAO Summary Report_YYYYMM (AverageDayWage_12Month, DailySalary, Days_AnnualLeave_LastMonth, Days_AnnualLeave, Days_StatutoryHolidays, Days_MaternityLeave, Days_SickLeave, Days_Paid Parental Time Off, NoPayLeaveCalculationBase, Days_NoPayLeave, Days_NoPayLeave_LastMonth etc.).
* Allowance plan report (Transportation Allowance).
* Calendar/Payroll period info (from `modCalendarService`).
* Shared leave splitting logic from Sub1 (we can reuse the same functions for counting days).

**Key items:**

1. **Base Pay / Base Pay (Temp) Check**

   * From Workforce Detail, using integer Monthly Salary (per rounding rule) and **Actual working days**; apply PDD formula (Base Pay = Actual working day * Monthly Salary / relevant divisor).
   * Round to 2 decimals (`RoundResult`).
   * Write to `Base Pay 60001000 Check` and `Base Pay(Temp) 60101000 Check`.

2. **Salary Adj / Transport Allowance Adj (Maternity + Paternity leave days)** 

   * Using new/unpaid **Maternity + Paternity Leave TOTAL DAYS** from Employee_Leave_Transactions_Report.
   * Use Monthly Salary (integer) and calendar days of last month.
   * Apply formulas: `- Monthly Salary / calendar days LastMonth * leave days`.
   * Transportation allowance adjustment uses Amount from Allowance plan report similarly.
   * Round via `RoundResult`.

3. **Transport Allowance Check**

   * Based on **Unpaid Leave** days in current month (from Employee_Leave_Transactions_Report) and Transportation Allowance from Allowance plan report.
   * Apply formula: `Amount - Amount / calendar days CurrentMonth * No pay leave days`.
   * Round to 2 decimals.

4. **Total EAO Adj Check**

   * From EAO Summary:

     * `days = Days_AnnualLeave_LastMonth + Days_AnnualLeave + Days_StatutoryHolidays`.
     * `Total EAO Adj = (AverageDayWage_12Month - DailySalary) * days`. 
   * Round to 2 decimals.

5. **Maternity Leave Payment / Sick Leave Payment / PPTO payment**

   * Use EAO Summary fields:

     * `DayWage_Maternity/Paternity/Sick Leave`, `Days_MaternityLeave`, `Days_SickLeave`, `Days_Paid Parental Time Off`.
   * Formulas per PDD:

     * Maternity: `DayWage * Days_MaternityLeave`.
     * Sick: `DayWage * Days_SickLeave`.
     * PPTO: `Max(DailySalary, AverageDayWage_12Month*80%) * Days_Paid Parental Time Off`. 
   * Round results and write to Check columns.

6. **No Pay Leave Deduction Check**

   * From EAO Summary: `NoPayLeaveCalculationBase * (Days_NoPayLeave + Days_NoPayLeave_LastMonth)`.
   * Round and map to Check.

7. **Untaken Annual Leave Payment Check**

   * Use EAO Summary + Workforce Detail (Monthly Salary integer).
   * `MAX(Monthly Salary / 22, AverageDayWage_12Month) * Untaken Annual Leave Days`. 

All repeated patterns (pull from EAO Summary, map WEIN, apply formula) should go through `modEAOService` helpers for maintainability.

---

### 6.5 Incentives & Variable Pay Checks

Module: `modCheckResult_Incentives`. Sources include One time payment report, 2025QX Payout Summary, Inspire Awards payroll report, RSU Dividend global/EY, 额外表, Merck Payroll Summary, etc. 

**Patterns:**

* Use `GroupByEmployeeAndType` from `modGroupingService` to ensure only one amount per employee per plan type.
* Map by Employee ID → WEIN.
* Round to 2 decimals (except where integer required, e.g. some tax gross‑ups).

**Items (not exhaustive but grouped):**

1. **AIP / Annual Incentive**

   * From One time payment report where `[Plan = AIP]`; group by Employee ID; map to `Annual Incentive 60201000 Check`.

2. **Sales Incentive (Quantitative / Qualitative)**

   * From `2025QX Payout Summary`:

     * `Pay Item = Sales Incentive Plan` → Quantitative.
     * `Pay Item = Qualitative Incentive Plan` → Qualitative. 

3. **Inspire Cash & Inspire Points**

   * From Inspire Awards Payroll Report:

     * Plan = `Inspire Cash` & `Inspire Points Value`.
   * Group by Employee ID, sum `Actual Payment – Amount`.

4. **Shares Dividend**

   * **May** – RSU Dividend global report;
   * **June** – RSU Dividend EY report;
   * Use `ExchangeRates` in config (`RateType = RSU_Global` / `RSU_EY`).
   * For each employee, `Gross Award Amount to be Paid × Exchange rate`, aggregated and mapped to `Shares Dividend Check`. 

5. **Red Packet, Lump Sum Bonus, Sign On Bonus, Retention Bonus, Referral Bonus, Manager of the Year Award, MD Award**

   * From One time payment report; Plan‑type filters as per PDD; grouped & mapped to their Check codes (same naming as Payroll Report chart of accounts).

6. **Flexible benefits / Other Allowance / Other Bonus / Other Rewards / Goods & Services Differential**

   * From `[特殊奖金]` sheet in 额外表; keyed by WEIN. 

7. **IA Pay Split Check**

   * Same as Sub1 but now for Check Result:

     * From Merck Payroll Summary: `IA Pay Split = Net Pay (include EAO & leave payment) + MPF EE MC + MPF EE VC`.
     * Map Employee ID → WEIN.

All these Check values are compared later against Payroll Report values in Diff columns.

---

### 6.6 Final Payment (Termination‑related) Checks

Module: `modCheckResult_FinalPayment`. Inputs: 额外表.[Final payment], Workforce Detail, Termination flex, Allowance plan, EAO Summary. 

**Key parts:**

1. **Severance Payment & Long Service Payment**

   * Using:

     * WEIN → Employee Code → Termination Date (flex Termination).
     * WEIN → Employee ID → Last Hire Date & Monthly Salary (rounded) from Workforce Detail.
     * 额外表.[Final payment]: `MSD policy or Statutory Policy`, `Termination Type`.
   * Calculate **Years of Service (YOS)**:

     * `YOS = (Termination Date – Last Hire Date) / 365` → rounded to 2 decimals.
   * Decision:

     * If TerminationType indicates **redundancy** or YOS < 5 → **Severance**.
     * Else (non‑redundancy & YOS ≥ 5) → **Long Service**.
   * Formula implementation:

     * For MSD policy: `Base Pay * Min(24, YOS)`.
     * For Statutory: `Min( Min(Monthly Salary * 2/3, 15000) * YOS, 390000)`. 

2. **Payment in lieu of notice (PIL EE to ER / ER to EE)**

   * From 额外表.[Final payment]: `PIL Indicator`, `Notice Given Date`, `Notice Period`.
   * Additional parameters:

     * `Monthly Base Pay Check`, `Monthly Base Pay(Temp) Check` (from Master Data block).
     * Transport allowance Amount from Allowance plan report.
     * `TotalWage_12Month` from EAO Summary.
     * Termination Date from flex Termination.
   * Compute:

     * `NoticeDays = Notice Period`.
     * `PIL EE to ER Days = NoticeDays − (TerminationDate − NoticeGivenDate)`.
     * `PIL ER to EE Days = NoticeDays − (TerminationDate − NoticeGivenDate)`.
   * Then apply formulas as per PDD (choose min / max between monthly wage based and TotalWage_12Month/12 etc.). 

3. **Gratuities & Back Pay**

   * Directly from 额外表.[Final payment].
   * Map WEIN and write to Check columns.

4. **Year End Bonus**

   * Termination case: from current terminations; use Monthly Salary (integer) for each terminated employee.
   * December case: if current month is December, compute:

     * If service < 1 year: `Monthly Salary / Annual Period * Service Period`.
     * Else: `Monthly Salary`.
   * Rely on `GetPayrollContext` & dates.

All formulas must be implemented exactly as in PDD, but the **parameters** (policy choice, thresholds) live in config / 额外表, not in code.

---

### 6.7 Contribution (MPF / ORSO) Checks

Module: `modCheckResult_Contribution`. Uses Payroll Report, Check Result itself, and 额外表.[MPF&ORSO]. 

**MPF Relevant Income Check**

* Build from multiple components (Base Pay, bonuses, leave payments, allowances, etc.) as listed in PDD’s big table; many are Check/Benchmark fields already computed.
* Add Goods & Services Differential (from 额外表.[特殊奖金]).
* Sum them into `MPF Relevant Income Check`.

**MPF VC Relevant Income Check**

* Similar but with slightly different composition per PDD (Base Pay, Base Pay(Temp), rental reimbursement, Total EAO Adj, Salary Adj, Maternity / Paternity / PPTO / Sick payments, minus No Pay Leave Deduction, etc.).

**MPF EE MC & MPF ER MC Check**

* Both: `MIN(MPF Relevant Income Check * 5%, 1500)`.

**MPF EE VC / MPF ER VC Check**

* MPF EE VC:

  * `MPF VC Relevant Income Check * MPF EE VC Percentage` (percentage from [MPF&ORSO] in 额外表).
* MPF ER VC:

  * `MPF VC Relevant Income Check * MPF ER VC Percentage − MPF ER MC`.
  * If result < 0 → 0; else apply rounding: `ROUND(MPF VC Relevant Income Check * MPF ER VC Percentage, 2) − MPF ER MC`. 

**ORSO Relevant Income / ORSO EE / ORSO ER / ORSO ER Adj / Percent of ORSO EE**

* ORSO Relevant Income Check:

  * Monthly Salary (integer) from Workforce Detail.
* ORSO EE:

  * `ORSO Relevant Income Check * 5%`.
* ORSO ER:

  * `ORSO Relevant Income Check * Percent Of ORSO ER` (from 额外表.[MPF&ORSO]).
* ORSO ER Adj, Percent Of ORSO EE:

  * Direct from [MPF&ORSO], mapped to Check.

All percentage values come from **参数表** – never hard‑coded.

---

### 6.8 Benefits for Tax – Inspire Points Gross‑up

Module: `modCheckResult_BenefitsTax`. 

* From Inspire Awards Payroll Report:

  * Plan = `Inspire Points Value`.
* Group by Employee ID; sum `Actual Payment – Amount`.
* Compute per PDD:

  * `GrossUpAmount = ROUNDUP(Actual Payment – Amount / (1 - 0.17) * 0.17, 0)`.
* Map Employee ID → WEIN; write to `Inspire Points (Gross Up) 60701000 Check`.

Note: here rounding is not 2 decimals but **`ROUNDUP(..., 0)`** (integer). This is an exception and must be implemented explicitly.

---

## 7. Diff Engine & Formatting

Module: `modDiffEngine` + `modFormattingService`.

### 7.1 Diff calculation

**Design**

* Use header metadata to know which columns pair:
  e.g. `Monthly Base Pay` Benchmark col X vs `Monthly Base Pay Check` col Y vs `Monthly Base Pay Diff` col Z.

**Steps:**

1. For each defined field pair:

   * For each data row:

     * Read Benchmark value and Check value.
     * Handle type conversions (dates vs numeric vs text).
     * If both blank → Diff = TRUE.
     * Else if special rule (e.g. Last Hired Date):

       * If Benchmark or Check < 2025‑01‑01 → Diff = TRUE regardless. 
     * Else:

       * Compare:

         * Dates: compare underlying `CLng(CDate(value))`.
         * Numbers: compare within small epsilon if needed (to avoid float glitches).
         * Text: case‑insensitive string compare after trimming.
       * If equal → Diff = TRUE; else FALSE.
2. Write result into Diff cell (TRUE/FALSE).

### 7.2 FALSE counting & red header highlight

Leverage the shared service requested earlier:

```vb
Public Sub ApplyDiffSummaryFormatting( _
    ws As Worksheet, _
    headerRow As Long, _
    firstDiffCol As Long, _
    lastDiffCol As Long, _
    firstDataRow As Long, _
    lastDataRow As Long)

    Dim c As Long, rng As Range, countFalse As Long
    For c = firstDiffCol To lastDiffCol
        Set rng = ws.Range(ws.Cells(firstDataRow, c), ws.Cells(lastDataRow, c))
        countFalse = WorksheetFunction.CountIf(rng, "FALSE")
        With ws.Cells(headerRow, c)
            .Value = countFalse
            If countFalse > 0 Then
                .Interior.Color = vbRed
                .Font.Color = vbWhite
            Else
                .Interior.ColorIndex = xlNone
            End If
        End With
    Next c
End Sub
```

This exactly implements requirement:

> Count FALSE per Diff column, write count in first row, and if >0, highlight cell in red. 

The header row & Diff column range will be looked up via template metadata; we do **not** hard‑code column letters.

---

## 8. Rounding Rules – Subprocess 2 Usage

We strictly reuse Sub1’s shared rounding service:

* **Monthly Salary**

  * Immediately rounded to the nearest whole number via `RoundMonthlySalary` when read for the first time. This applies to all checks where Monthly Salary is used (e.g. Base Pay, Year End Bonus, Severance, Long Service, ORSO Relevant Income).
* **All calculation results**

  * Use `RoundResult(..., 2 decimals)` before writing to Check columns, unless PDD explicitly defines a different rounding (e.g. `ROUNDUP(..., 0)` for Inspire Points Gross Up, or integer caps for MPF).
* MPF / ORSO contributions:

  * Ensure we round after min/max logic as specified (some formulas have a `MIN(...,1500)` before rounding).

Centralising this in `modRoundingService` prevents inconsistent rounding behaviours across fields.

---

## 9. Error Handling, Logging & Testing

### 9.1 Error handling

* All public entry subs (`SP2_Run`, group subs) have `On Error GoTo` blocks.
* Use `LogError` helper to append:

  * Procedure, workbook, sheet, row, WEIN (if available), error number & description.
* Store logs in a `Log_SP2` sheet in macro workbook and optionally export to `.csv` for PAD to attach to notifications.

### 9.2 Idempotency

* Sub2 is naturally **idempotent**: each run rebuilds `HK Payroll Validation Output YYYYMMDD.xlsx` from scratch.
* There is no need to track processed leaves here (Sub1 already does “new/unpaid” in Attendance/VariablePay), but we re‑use the same logic when Sub2 needs “new unpaid” for Salary Adj and Transport Allowance Adj.

### 9.3 Testing strategy

* **Unit tests** for:

  * Mapping WEIN ⇔ Employee ID ⇔ Employee Code.
  * EAO functions and MPF/ORSO calculations.
  * Diff engine (including special rules for Last Hired Date).
* **Scenario tests**:

  * One small test month with synthetic data for each pay item & final payment scenario.
  * Edge cases: cross‑year dates, pre‑2025 Last Hired Date, zero or negative MPF ER VC, Termination with short YOS, etc.
* **Regression baseline**:

  * Run Sub2 against one full real month and compare output with manually confirmed validation files.

---

## 10. Summary – Alignment with Requirements

* **Configurable schedule and business‑day logic**

  * Reused `PayrollCalendar` & HK holiday services from Sub1; Sub2’s HC checks, PIL and headcount all rely on this configuration instead of hard‑coded dates.

* **Centralised grouping & summing**

  * All repeating patterns where “If duplicate records exist, sum first” (One time payment, Inspire, SIP/2025QX, RSU, leave EAO, etc.) are implemented via `GroupByEmployeeAndType + SumAmount` in `modGroupingService`.

* **No hard‑coded rates or policy logic**

  * Exchange rates, MPF/ORSO percentages, KGG/MSD vs Statutory policy, final payment rules, and special bonuses come from `config.xlsx` and `额外表`, not from code.

* **Formatting logic (Diff summary)**

  * Implemented as a shared VBA service used **only in Sub2**, counting FALSE per Diff column and highlighting headers in red when needed.

* **Rounding rules enforced consistently**

  * Monthly salary always rounded to whole number at source.
  * All calculation outputs rounded to 2 decimals (with explicit exceptions where PDD mandates different rounding).
  * All Sub2 calculations call into `modRoundingService`.

With this plan, Subprocess 2 becomes a **config‑driven, reusable, and verifiable** validation engine layered on top of the shared services from Subprocess 1, fully aligned to the detailed PDD and the technology constraints (PAD + VBA, no Power Query).
