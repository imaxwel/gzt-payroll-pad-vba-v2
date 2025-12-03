Here’s a pattern that lets **Subprocess 1 and Subprocess 2 run different macros, but share global state and shared services cleanly** in a single VBA project – without turning the codebase into spaghetti. 

---

# 1. Design Goals

* **Two separate entry macros** (for PAD):

  * `Run_Subprocess1`
  * `Run_Subprocess2`
* **Shared “global” context**:

  * Run parameters (input/output paths, payroll month, run date).
  * Payroll calendar / cut‑off info.
  * Common mappings (WEIN ↔ Employee ID, etc.).
* **Shared service functions**:

  * Config & parameter reading.
  * Rounding rules.
  * Grouping “sum per employee per type”.
  * Calendar / HK holiday logic.
  * EAO / leave helpers.
* **Avoid fragile globals**:

  * One structured global context object instead of many scattered `Public` variables.
  * All global data initialised through a single `InitAppContext` function.

> Important reality check: **VBA global variables only exist per Excel instance**.
> If PAD opens a *new* Excel process for Subprocess 2, it will not see globals created by Subprocess 1.
> So: each entry macro must call `InitAppContext` and be self‑sufficient.

---

# 2. Recommended Module Layout

## 2.1 Entry & Context

**Standard modules (BAS):**

1. `modEntryPoints` – PAD entry macros.
2. `modAppContext` – global types and the single global `AppContext`.

## 2.2 Shared Services (Sub1 + Sub2)

3. `modConfigService` – config, file mapping, payroll schedule.
4. `modRoundingService` – monthly salary & 2‑decimal rounding.
5. `modAggregationService` – `GroupByEmployeeAndType + SumAmount`.
6. `modCalendarService` – HK business days, public holidays, cross‑month splits.
7. `modEmployeeMappingService` – WEIN / Employee ID / Employee Code dictionaries.
8. `modEAOService` – EAO / leave and cross‑month splitting helpers.
9. `modLoggingService` – logging to sheet / text file.
10. `modFormattingService` – generic formatting (e.g. Diff FALSE summary row).

## 2.3 Subprocess‑specific Orchestration

11. `modSP1_Main` – orchestration for Subprocess 1.

12. `modSP1_Attendance` – all Attendance sheet logic.

13. `modSP1_VariablePay` – all VariablePay sheet logic.

14. `modSP2_Main` – orchestration for Subprocess 2.

15. `modSP2_CheckResult_MasterData`

16. `modSP2_CheckResult_PayItems`

17. `modSP2_CheckResult_Contributions`

18. `modSP2_CheckResult_FinalPayments`

19. `modSP2_CheckResult_TaxBenefits`

20. `modSP2_CheckResult_Diff`

21. `modSP2_HCCheck`

You can combine some of these if you want fewer modules, but **keep shared vs SP1‑specific vs SP2‑specific clearly separated**.

---

# 3. Global Context Pattern

## 3.1 Type Definitions (modAppContext)

```vb
'==== modAppContext ====
Option Explicit

' Parameters passed in / derived at run-time
Public Type tRunParams
    InputFolder As String
    OutputFolder As String
    ConfigFolder As String
    PayrollMonth As String  ' "YYYYMM"
    RunDate As Date
End Type

' Per-month payroll calendar, from config
Public Type tPayrollContext
    PayrollMonth As String
    MonthStart As Date
    MonthEnd As Date
    PrevMonthStart As Date
    PrevMonthEnd As Date
    PayDate As Date
    PreviousCutoff As Date
    CurrentCutoff As Date
End Type

' Application-wide shared state
Public Type tAppContext
    RunParams As tRunParams
    Payroll As tPayrollContext
    ' Common mappings
    DictWeinToEmpId As Object   ' Scripting.Dictionary
    DictEmpIdToWein As Object   ' Scripting.Dictionary
    DictEmpCodeToWein As Object ' optional
    ' Other shared caches
    IsInitialised As Boolean
End Type

Public G As tAppContext   ' <== the SINGLE global variable
```

## 3.2 Initialise & Reset Context

```vb
Public Sub InitAppContext(p As tRunParams)
    ' Always reset first
    ResetAppContext
    
    ' Store run parameters
    G.RunParams = p
    
    ' Load payroll calendar from config
    G.Payroll = GetPayrollContext(p.PayrollMonth)  ' from modConfigService
    
    ' Build shared mappings
    BuildEmployeeMappings G   ' from modEmployeeMappingService
    
    G.IsInitialised = True
End Sub

Public Sub ResetAppContext()
    G.IsInitialised = False
    Set G.DictWeinToEmpId = Nothing
    Set G.DictEmpIdToWein = Nothing
    Set G.DictEmpCodeToWein = Nothing
    ' If you keep workbook references here, set them to Nothing too
End Sub
```

> **Best practice**: *all* modules that use `G` should check `If Not G.IsInitialised Then InitAppContext ...` if there’s any chance they might be called standalone.

---

# 4. Entry Macros for Subprocess 1 & 2

## 4.1 Entry Module (modEntryPoints)

```vb
'==== modEntryPoints ====
Option Explicit

Public Sub Run_Subprocess1()
    Dim p As tRunParams
    
    ' 1) Fetch parameters (from a config sheet or named ranges)
    p = LoadRunParamsFromWorkbook()  ' from modConfigService
    
    ' 2) Initialise global context
    InitAppContext p
    
    ' 3) Execute Subprocess 1 logic
    SP1_Execute
End Sub

Public Sub Run_Subprocess2()
    Dim p As tRunParams
    
    p = LoadRunParamsFromWorkbook()
    InitAppContext p
    
    SP2_Execute
End Sub
```

PAD just calls `Run_Subprocess1` or `Run_Subprocess2`. Both:

* Initialise the same global `G`.
* Use the same shared services (config, rounding, grouping, EAO, etc.).
* Run separate orchestration pipelines.

---

# 5. Shared Service Modules

Below is the **shape** of key shared modules. They are all `Option Explicit`, and most functions are `Public` so they can be called from both SP1 and SP2 modules.

## 5.1 Config & Path Service (modConfigService)

```vb
'==== modConfigService ====
Option Explicit

Public Function LoadRunParamsFromWorkbook() As tRunParams
    Dim p As tRunParams
    With ThisWorkbook.Worksheets("Runtime")
        p.InputFolder = .Range("InputFolder").Value
        p.OutputFolder = .Range("OutputFolder").Value
        p.ConfigFolder = .Range("ConfigFolder").Value
        p.PayrollMonth = .Range("PayrollMonth").Value
        p.RunDate = .Range("RunDate").Value
    End With
    LoadRunParamsFromWorkbook = p
End Function

Public Function GetPayrollContext(payrollMonth As String) As tPayrollContext
    Dim ctx As tPayrollContext
    ' Read from config.xlsx, PayrollCalendar sheet
    ' Assign ctx.MonthStart, ctx.MonthEnd, ctx.PayDate, etc.
    GetPayrollContext = ctx
End Function

Public Function GetExchangeRate(rateName As String) As Double
    ' Read from config ExchangeRates table
End Function

Public Function GetInputFilePath(logicalName As String) As String
    ' Maps "PayrollReport", "WorkforceDetail", etc. to concrete file paths.
End Function
```

## 5.2 Rounding (modRoundingService)

```vb
'==== modRoundingService ====
Option Explicit

Public Function RoundMonthlySalary(ByVal v As Variant) As Double
    If IsNumeric(v) Then
        RoundMonthlySalary = WorksheetFunction.Round(CDbl(v), 0)
    Else
        RoundMonthlySalary = 0
    End If
End Function

Public Function RoundAmount2(ByVal v As Variant) As Double
    If IsNumeric(v) Then
        RoundAmount2 = WorksheetFunction.Round(CDbl(v), 2)
    Else
        RoundAmount2 = 0
    End If
End Function
```

All SP1 and SP2 calculations **must** go through these; don’t hand‑sprinkle `Round` everywhere.

## 5.3 Aggregation (modAggregationService)

```vb
'==== modAggregationService ====
Option Explicit

Public Function GroupByEmployeeAndType( _
    ByVal dataRange As Range, _
    ByVal employeeColName As String, _
    ByVal typeColName As String, _
    ByVal amountColName As String _
) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Find column indices by header names...
    ' Loop rows, build key employee|type, accumulate amount (Double)
    
    Set GroupByEmployeeAndType = dict
End Function
```

Any “if multiple records exist, sum first” rule (Sub1 **and** Sub2) uses this.

## 5.4 Calendar & HK Holidays (modCalendarService)

```vb
'==== modCalendarService ====
Option Explicit

Public Function IsWeekend(d As Date) As Boolean
    IsWeekend = (Weekday(d, vbMonday) > 5)
End Function

Public Function IsHKPublicHoliday(d As Date) As Boolean
    ' Lookup in config calendar table
End Function

Public Function IsBusinessDay(d As Date) As Boolean
    IsBusinessDay = Not IsWeekend(d) And Not IsHKPublicHoliday(d)
End Function

Public Sub SplitByMonth( _
        ByVal startDate As Date, _
        ByVal endDate As Date, _
        ByRef spans As Collection)
    ' Populate spans with segments per calendar month
End Sub
```

Used in both Sub1 (Attendance/VariablePay) and Sub2 (EAO / service periods / YOS).

## 5.5 WEIN ↔ Employee Mapping (modEmployeeMappingService)

```vb
'==== modEmployeeMappingService ====
Option Explicit

Public Sub BuildEmployeeMappings(ByRef ctx As tAppContext)
    Dim wb As Workbook, ws As Worksheet, lastRow As Long
    Dim empId As String, wein As String, empCode As String
    
    Set ctx.DictWeinToEmpId = CreateObject("Scripting.Dictionary")
    Set ctx.DictEmpIdToWein = CreateObject("Scripting.Dictionary")
    Set ctx.DictEmpCodeToWein = CreateObject("Scripting.Dictionary")
    
    Set wb = Workbooks.Open(GetInputFilePath("WorkforceDetail"))
    Set ws = wb.Worksheets(1) ' or by name
    
    ' Loop Workforce Detail rows, populate dictionaries
    
    wb.Close SaveChanges:=False
End Sub

Public Function WeinFromEmpId(empId As String) As String
    WeinFromEmpId = ""
    If Not G.DictEmpIdToWein Is Nothing Then
        If G.DictEmpIdToWein.Exists(empId) Then _
            WeinFromEmpId = G.DictEmpIdToWein(empId)
    End If
End Function

Public Function EmpIdFromWein(wein As String) As String
    ' Similar pattern
End Function
```

All SP1/SP2 logic that needs to convert between IDs uses these central helpers.

## 5.6 EAO & Leave (modEAOService)

This holds helpers such as:

* `GetEAORecord(wein, monthKey)`.
* `CalcAnnualLeaveEAOAdj(...)`.
* `CalcSickLeaveEAOAdj(...)`.
* `CalcNoPayLeaveDeduction(...)`, etc.

Each function is pure (only inputs/outputs), uses `RoundMonthlySalary` & `RoundAmount2`, and is used by both Sub1 and Sub2.

## 5.7 Diff Summary Formatting (modFormattingService)

```vb
'==== modFormattingService ====
Option Explicit

Public Sub SummarizeDiffColumns( _
        ByVal ws As Worksheet, _
        ByVal headerRow As Long, _
        ByVal firstDataRow As Long, _
        ByVal lastDataRow As Long, _
        ByVal firstDiffCol As Long, _
        ByVal lastDiffCol As Long)

    Dim col As Long, falseCount As Long, rng As Range
    
    For col = firstDiffCol To lastDiffCol
        Set rng = ws.Range(ws.Cells(firstDataRow, col), ws.Cells(lastDataRow, col))
        falseCount = WorksheetFunction.CountIf(rng, "FALSE")
        
        With ws.Cells(headerRow, col)
            .Value = falseCount
            If falseCount > 0 Then
                .Interior.Color = vbRed
                .Font.Color = vbWhite
            Else
                .Interior.ColorIndex = xlColorIndexNone
            End If
        End With
    Next col
End Sub
```

Sub2 uses this on `Check Result` after Diff columns are computed.

---

# 6. Subprocess 1 Orchestration (modSP1_Main)

```vb
'==== modSP1_Main ====
Option Explicit

Public Sub SP1_Execute()
    If Not G.IsInitialised Then
        Err.Raise vbObjectError + 1000, "SP1_Execute", "AppContext not initialised"
    End If
    
    ' 1. Create Flexi form output workbook
    Dim wbFlex As Workbook
    Set wbFlex = CreateFlexiOutputWorkbook(G.RunParams.OutputFolder, G.Payroll)
    
    ' 2. Load raw flexiform data into NewHire/InformationChange/...
    SP1_LoadFlexiformData wbFlex
    
    ' 3. Populate Attendance (leave splitting etc.)
    SP1_PopulateAttendance wbFlex
    
    ' 4. Populate VariablePay (One time payment, SIP, RSU, etc.)
    SP1_PopulateVariablePay wbFlex
    
    ' 5. Final formatting & save
    SP1_FinalizeFlexOutput wbFlex
End Sub
```

SP1 helper subs live in:

* `modSP1_Attendance`
* `modSP1_VariablePay`

but they freely call shared services (`modAggregationService`, `modEAOService`, etc.) via `Public` functions.

---

# 7. Subprocess 2 Orchestration (modSP2_Main)

```vb
'==== modSP2_Main ====
Option Explicit

Public Sub SP2_Execute()
    If Not G.IsInitialised Then
        Err.Raise vbObjectError + 1001, "SP2_Execute", "AppContext not initialised"
    End If
    
    Dim wbOut As Workbook
    Set wbOut = CreateValidationOutputWorkbook(G.RunParams.OutputFolder, G.Payroll)
    
    ' 1. Copy Payroll Report to Check Result (Benchmark Data)
    SP2_BuildBenchmark wbOut
    
    ' 2. Populate Check columns by group
    SP2_Check_MasterData wbOut
    SP2_Check_PayItems wbOut
    SP2_Check_Contributions wbOut
    SP2_Check_FinalPayments wbOut
    SP2_Check_TaxBenefits wbOut
    
    ' 3. Compute Diff columns
    SP2_ComputeDiff wbOut
    
    ' 4. Diff summary row
    SP2_SummarizeDiff wbOut   ' wraps SummarizeDiffColumns
    
    ' 5. HC Check sheet
    SP2_BuildHCCheck wbOut
    
    wbOut.Save
End Sub
```

Again, all the shared logic (rounding, EAO calcs, grouping, mapping) comes from shared modules; Sub2 modules just orchestrate.

---

# 8. How “Shared Globals” Actually Work

* Both `Run_Subprocess1` and `Run_Subprocess2` **call `InitAppContext`** with a `tRunParams` instance.
* `InitAppContext` fills the single global `G` and builds shared dictionaries.
* All shared services either:

  * Use `G` directly (e.g., `GetInputFilePath` can read `G.RunParams.ConfigFolder`), or
  * Are pure functions that accept parameters and don’t touch globals.

**Rules to keep this safe:**

1. **Single global struct**

   * Only `G As tAppContext` is global.
   * Avoid scattering `Public` scalar globals (`Public InputFolder As String`, etc.).

2. **Explicit initialisation**

   * No function should assume `G` is initialised; SP1/SP2 orchestration always initialises before use.
   * If you call SP1/2 helpers manually during development, you can call `InitAppContext` first.

3. **No cross‑subprocess assumptions**

   * Subprocess 2 should *not* assume Subprocess 1 has already been run in the same Excel session, even though it might be in practice.
   * Each `Run_SubprocessX` is a complete, self‑contained run.

4. **Option Explicit everywhere**

   * Prevents accidental creation of “ghost globals” due to typos.

---

# 9. Summary

* Use **different entry macros** (`Run_Subprocess1`, `Run_Subprocess2`) for PAD.
* Use a **single global context (`G As tAppContext`)** in `modAppContext` to hold:

  * Run parameters,
  * Payroll calendar,
  * Shared dictionaries and caches.
* Initialise that context via `InitAppContext` from both Subprocess entrypoints.
* Place all **shared services** (config, rounding, grouping, calendar, EAO, mapping, logging, formatting) in dedicated standard modules with `Public` functions.
* Keep Subprocess 1 and Subprocess 2 orchestration in their own modules, **only consuming** the shared services and shared global context.

This gives you:

* Separation of concerns,
* Shared logic in one place,
* Predictable “global” behaviour within an Excel session,
* And clear, PAD‑friendly macros for each subprocess.
