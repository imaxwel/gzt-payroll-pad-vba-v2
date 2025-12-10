# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is the **HK Payroll Input & Validation** automation system - a VBA-based Excel application for processing monthly payroll data. The system consists of two subprocess pipelines:

- **Subprocess 1**: Generates `Flexi form out put YYYYMMDD.xlsx` with employee data including NewHire, InformationChange, SalaryChange, Termination, Attendance, and VariablePay sheets
- **Subprocess 2**: Generates `HK Payroll Validation Output YYYYMMDD.xlsx` with validation results comparing benchmark data against calculated values

The system is orchestrated via **Power Automate Desktop (PAD)** and processes 15+ Excel input files monthly.

## Technology Stack

- **Language**: VBA (Visual Basic for Applications) with late-bound COM objects
- **Platform**: Excel 2016+ on Windows 10/11
- **Orchestration**: Power Automate Desktop (PAD)
- **No dependencies on**: Power Query, third-party libraries, or external DLLs

## Architecture

### Global Context Pattern

The codebase uses a single global context object (`G As tAppContext` in `modAppContext.bas`) to share state across all modules:

```vb
Public Type tAppContext
    RunParams As tRunParams          ' Input/output paths, payroll month, run date
    Payroll As tPayrollContext       ' Calendar dates, cutoffs, pay date
    DictWeinToEmpId As Object        ' Employee ID mappings
    DictEmpIdToWein As Object
    DictEmpCodeToWein As Object
    DictWeinToEmpCode As Object
    configWb As Workbook             ' Config workbook reference
    ExtraTableWb As Workbook         ' 额外表 workbook
    IsInitialised As Boolean
End Type
```

Both subprocesses initialize this context via `InitAppContext` at startup.

### Module Organization

**Entry Points & Context:**
- `modEntryPoints.bas` - PAD entry macros: `Run_Subprocess1`, `Run_Subprocess2`
- `modAppContext.bas` - Global types and the single `G` variable

**Shared Services (used by both subprocesses):**
- `modConfigService.bas` - Config reading, file paths, payroll schedule
- `modRoundingService.bas` - Consistent rounding rules (monthly salary, 2-decimal amounts)
- `modAggregationService.bas` - "Group by employee and type, sum amounts" logic
- `modCalendarService.bas` - HK business days, holidays, cross-month date splitting
- `modEmployeeMappingService.bas` - WEIN ↔ Employee ID ↔ Employee Code conversions
- `modEAOService.bas` - EAO (Estimated Annual Obligation) calculations and leave helpers
- `modLoggingService.bas` - Logging to sheet and text file
- `modFormattingService.bas` - Generic output formatting (e.g., Diff FALSE summary row)

**Subprocess 1 Modules:**
- `modSP1_Main.bas` - Orchestration for Subprocess 1
- `modSP1_Attendance.bas` - Attendance sheet logic (leave days, cross-month splits)
- `modSP1_VariablePay.bas` - Variable pay sheet logic (bonuses, RSU, AIP, etc.)

**Subprocess 2 Modules:**
- `modSP2_Main.bas` - Orchestration for Subprocess 2
- `modSP2_CheckResult_MasterData.bas` - Master data validation checks
- `modSP2_CheckResult_PayItems.bas` - Pay items validation
- `modSP2_CheckResult_Contribution.bas` - MPF/ORSO contribution validation
- `modSP2_CheckResult_FinalPayment.bas` - Final payment (termination) validation
- `modSP2_CheckResult_BenefitsTax.bas` - Benefits and tax validation
- `modSP2_CheckResult_Incentives.bas` - Incentive validation
- `modSP2_CheckResult_Diff.bas` - Diff column computation (Benchmark vs Check)
- `modSP2_HCCheck.bas` - Headcount check sheet generation

### Key Design Principles

1. **Late-bound COM**: All objects use `Object` type (not `Scripting.Dictionary`, `ADODB.*`), no early binding
2. **Single global context**: Only `G As tAppContext` is global; avoid scattered `Public` variables
3. **No hard-coded paths/rules**: All configuration comes from `config.xlsx` and `额外表.xlsx`
4. **Consistent rounding**: All calculations must use `RoundMonthlySalary()` or `RoundAmount2()` from `modRoundingService`
5. **Error handling**: Every `Public Sub/Function` should have `On Error GoTo ErrHandler` with `LogError` calls

## Directory Structure

```
4codex/
├── code/                           # VBA .bas module files (source of truth)
│   ├── modAppContext.bas           # Global context types
│   ├── modEntryPoints.bas          # PAD entry macros
│   ├── mod*Service.bas             # Shared services
│   ├── modSP1_*.bas                # Subprocess 1 modules
│   └── modSP2_*.bas                # Subprocess 2 modules
├── config/                         # Configuration workbooks
│   └── config.xlsx                 # PayrollSchedule, Calendar, ExchangeRates sheets
├── Subprocess 1/
│   ├── input/                      # SP1 input files (15+ monthly Excel files)
│   └── output/                     # Flexi form out put YYYYMMDD.xlsx
├── Subprocess 2/
│   ├── input/                      # SP2 input files (additional validation sources)
│   └── output/                     # HK Payroll Validation Output YYYYMMDD.xlsx
├── doc/                            # User documentation (Chinese + English)
├── guides/                         # Technical implementation guides
└── HK_Payroll_Automation.xlsm      # Main control workbook (not in repo)
```

## Configuration Files

### config.xlsx

Located in `config/`, contains:
- **PayrollSchedule**: Monthly calendar (PayrollMonth, CutoffDate, PayDate, IsAIPMonth, IsRSUDivMonth, IsFlexBenefitMonth flags)
- **Calendar**: HK public holidays (Date, IsHKHoliday)
- **ExchangeRates**: FX rates (RSU_Global, RSU_EY, DefaultFX)

### 额外表.xlsx

Located in `Subprocess 1/input/` and `Subprocess 2/input/`, contains:
- **[需要每月维护]**: Monthly maintenance params (WEIN, PPTO EAO Rate input)
- **[MPF&ORSO]**: MPF/ORSO contribution percentages by employee
- **[特殊奖金]**: Special bonuses and allowances by employee
- **[Final payment]**: Termination payment parameters

**Important**: 额外表 must be updated monthly by HR team.

## Development Commands

### Setting Up the Main Workbook

The main control workbook `HK_Payroll_Automation.xlsm` is not in the repository. To create it:

1. Create a new Excel workbook and save as `HK_Payroll_Automation.xlsm` (macro-enabled)
2. Import all `.bas` files from `code/` folder via VBA Editor (Alt+F11 → File → Import)
3. Create a `Runtime` sheet with named ranges:
   - `InputFolder` (e.g., `C:\HK_Payroll\Subprocess 1\input\`)
   - `OutputFolder` (e.g., `C:\HK_Payroll\Subprocess 1\output\`)
   - `ConfigFolder` (e.g., `C:\HK_Payroll\config\`)
   - `PayrollMonth` (YYYYMM format, e.g., `202501`)
   - `RunDate` (date value)
   - `LogFolder` (e.g., `C:\HK_Payroll\log\`)
   - `SP_Status` (output: "OK" or "ERROR")
   - `SP_Message` (output: status message)

### Running Subprocesses Manually

**From Excel VBA:**
1. Open `HK_Payroll_Automation.xlsm`
2. Ensure `Runtime` sheet parameters are set correctly
3. Press Alt+F8 → Select `Run_Subprocess1` or `Run_Subprocess2` → Run

**From PAD (Production):**
- PAD opens the workbook, writes parameters to `Runtime` sheet, and executes the macro
- PAD reads `SP_Status` and `SP_Message` for result handling

### Debugging VBA Code

1. Open VBA Editor (Alt+F11)
2. Set breakpoints in relevant modules
3. Run macro via F5 or Alt+F8
4. Check `Log` sheet in control workbook for detailed execution logs
5. Check log files in `LogFolder` for persistent logs

### Extracting VBA Modules to .bas Files

When modifying VBA in Excel, export modules back to `code/` folder:
1. In VBA Editor, right-click module → Export File
2. Save to `code/` folder, overwriting existing `.bas` file
3. Commit changes to git

### Importing VBA Modules from .bas Files

To update `HK_Payroll_Automation.xlsm` with latest code:
1. In VBA Editor, right-click project → Import File
2. Select updated `.bas` file from `code/` folder
3. If module already exists, remove old one first

## Common Development Patterns

### Adding a New Shared Service Function

1. Add function to appropriate service module (e.g., `modCalendarService.bas`)
2. Make it `Public` so both subprocesses can call it
3. Use `On Error GoTo ErrHandler` with `LogError`
4. Call `EnsureInitialised` if function depends on `G` context
5. Use `RoundMonthlySalary` or `RoundAmount2` for numeric results
6. Export module to `code/` folder when done

Example:
```vb
Public Function MyNewHelper(param As String) As Double
    On Error GoTo ErrHandler

    EnsureInitialised

    ' Your logic here
    MyNewHelper = RoundAmount2(result)
    Exit Function

ErrHandler:
    LogError "modConfigService", "MyNewHelper", Err.Number, Err.Description
    Err.Raise Err.Number, "MyNewHelper", Err.Description
End Function
```

### Adding a New Input File

1. Update `GetInputFilePath()` in `modConfigService.bas` with new logical name
2. Update input file list in documentation (`doc/Configuration_Guide.md`)
3. Add file processing logic to relevant subprocess module

### Modifying Business Rules

1. Check if rule should be in config (preferred) or code
2. If config: update `config.xlsx` or `额外表.xlsx` structure and add loading logic
3. If code: modify relevant calculation function and update tests
4. Ensure rounding consistency via service functions

### Adding New Columns to Output

1. Update relevant sheet generation in `modSP1_Main.bas` or `modSP2_Main.bas`
2. Add column headers during sheet creation
3. Add data population logic
4. Update any formatting that depends on column positions

## Adaptive Header Detection Refactoring Plan

### Background

Many input Excel files have header rows that are NOT at row 1. The current VBA code has hardcoded assumptions (`ws.Rows(1)`, `ws.Cells(1, col)`) that fail when headers are not in the first row.

**Input files with non-standard header positions:**

| File | Sheet | Header Row | Key Column |
|------|-------|------------|------------|
| Workforce Detail - Payroll-AP.xlsx | Sheet1 | **Row 16** | Employee ID |
| Additional table.xlsx | 特殊奖金 | **Row 3** | WEIN |
| Additional table.xlsx | Previous Month Terminated HC | **Row 3** | WEIN |
| Allowance plan report.xlsx | Sheet1 | **Row 10** | Employee ID |
| AIP Payouts Payroll Report.xlsx | Sheet1 | **Row 7** | Employee WIN |
| RSU Dividend global report.xlsx | Cash Pay Instructions | **Row 13** | Employee Reference |
| Dividend EY report.xlsx | Dividend Payment Details | **Row 1** | Employee Number |

### Core Pattern: Adaptive Header Detection

The codebase already has `FindHeaderRow()` in `modAggregationService.bas`. All input file reading functions must use this pattern:

```vb
' Standard adaptive header detection pattern
headerRow = FindHeaderRow(ws, "Employee ID,EmployeeID,WEIN,WIN", 50)
If headerRow = 0 Then
    LogWarning "Module", "Function", "Could not find header row, defaulting to row 1"
    headerRow = 1
End If

' Build header index from detected row
Set headers = CreateObject("Scripting.Dictionary")
For c = 1 To ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    headers(UCase(Trim(CStr(Nz(ws.Cells(headerRow, c).Value, ""))))) = c
Next c

' Data starts from headerRow + 1
lastRow = ws.Cells(ws.Rows.Count, empCol).End(xlUp).Row
For i = headerRow + 1 To lastRow
    ' Process data rows
Next i
```

### Modules Requiring Refactoring

**High Priority (Multiple hardcoded row 1 assumptions):**

1. **modSP1_Main.bas**
   - `CopyFlexiformSheet()` - Lines ~166, 198, 220
   - `AddVariablePayColumns()` - Uses `ws.Rows(1)` directly

2. **modSP1_VariablePay.bas**
   - `ProcessOneTimePayment()` - Lines ~76-77, 140
   - `ProcessInspireAwards()` - Lines ~183-184
   - `ProcessAIPPayouts()` - Lines ~247-248 (needs Row 7 detection)
   - `ProcessRSUDividendGlobal()` - Lines ~738-742 (needs Row 13 detection)
   - `ProcessExtraTableSheet()` - needs Row 3 detection for 特殊奖金
   - `GetOrAddRow()` - Line ~899

3. **modSP1_Attendance.bas**
   - `LoadLeaveTransactions()` - Lines ~106-107 (currently hardcoded row 1)

**Medium Priority:**

4. **modSP2_CheckResult_MasterData.bas**
   - `LoadAllowanceData()` - Lines ~218-222 (needs Row 10 detection)

5. **modSP2_HCCheck.bas**
   - `CalculateExtraTableHC()` - Lines ~153-154

**Already Refactored (Reference):**

- `modSP2_CheckResult_Incentives.bas` - `ProcessOneTimePaymentCheck()`, `ProcessInspireCheck()` (Commit 1415661)
- `modSP2_CheckResult_MasterData.bas` - `LoadWorkforceData()` (searches 50 rows, 200 columns)
- `modSP1_Attendance.bas` - `LoadWorkforceHireDates()` (searches 50 rows, 200 columns)

### Implementation Steps

1. **Enhance FindHeaderRow() in modAggregationService.bas**
   - Increase default maxRows to 50
   - Add special keyword variants for each file type
   - Add logging when header is found at non-row-1 position

2. **Add file-specific keyword constants to modEmployeeMappingService.bas**
   ```vb
   ' Additional keyword variants for specific files
   Public Const AIP_HEADER_KEYWORDS As String = "Employee WIN,Worker"
   Public Const RSU_GLOBAL_KEYWORDS As String = "Employee Reference,Employee Number"
   Public Const ALLOWANCE_KEYWORDS As String = "Employee ID,EmployeeID"
   ```

3. **Refactor each module in priority order**
   - Replace `ws.Rows(1)` with `ws.Rows(headerRow)`
   - Replace `ws.Cells(1, col)` with `ws.Cells(headerRow, col)`
   - Add `FindHeaderRow()` call before any header-dependent operations
   - Update data loop: `For i = 2 To lastRow` → `For i = headerRow + 1 To lastRow`

4. **Testing with Python scripts**
   ```python
   # Verification script pattern
   import openpyxl

   def verify_header_detection(file_path, expected_keywords, expected_row):
       wb = openpyxl.load_workbook(file_path, read_only=True)
       ws = wb.active
       # Find header row using same logic as VBA
       for row in range(1, 51):
           for col in range(1, 201):
               val = ws.cell(row=row, column=col).value
               if val and str(val).upper() in [k.upper() for k in expected_keywords]:
                   assert row == expected_row, f"Expected row {expected_row}, found {row}"
                   return True
       return False
   ```

### Output Sheet Handling

**Note:** Output sheets (generated by the system) have a fixed structure with headers at row 4 (rows 1-3 contain metadata). These should continue to use hardcoded row numbers since we control the output format:

- `modSP2_CheckResult_FinalPayment.bas` - Row 4
- `modSP2_CheckResult_Diff.bas` - Row 4
- `modSP2_Main.bas` - Row 4

Consider defining a constant: `Const OUTPUT_HEADER_ROW As Long = 4`

## Known Issues & Workarounds

### Issue #94: Annual Leave Processing Error
- **Symptom**: `[modSP1_Attendance.ProcessAnnualLeave] #94: 无效使用 Null`
- **Context**: Related to null handling in EAO rate lookups
- **Status**: Investigation ongoing

### VBA Encoding Issues
- Some file names (e.g., `额外表.xlsx`) display as garbled text in VBA Editor on certain systems
- Use logical names in code via `GetInputFilePath()` instead of literal file names

### Excel Performance
- Processing 15+ input files with 1000+ employees can take 5-10 minutes
- Use `Application.ScreenUpdating = False` at start of long operations
- Use `Application.Calculation = xlCalculationManual` when writing large data ranges

## Testing Approach

### Manual Testing Checklist

**Subprocess 1:**
1. Place all required input files in `Subprocess 1/input/` folder
2. Set `PayrollMonth` to test month (e.g., `202501`)
3. Run `Run_Subprocess1`
4. Verify output file created: `Flexi form out put YYYYMMDD.xlsx`
5. Check each sheet has correct headers and data
6. Verify Attendance sheet has leave day calculations
7. Verify VariablePay sheet has all pay types aggregated

**Subprocess 2:**
1. Ensure Subprocess 1 completed successfully
2. Copy additional SP2 input files to `Subprocess 2/input/` folder
3. Run `Run_Subprocess2`
4. Verify output file: `HK Payroll Validation Output YYYYMMDD.xlsx`
5. Check "Check Result" sheet has Benchmark, Check, and Diff columns populated
6. Verify Diff summary row shows FALSE counts (highlighted red if > 0)
7. Check "HC Check" sheet has headcount reconciliation

### Data Validation

- **Attendance**: Check that cross-month leaves are split correctly by calendar month
- **VariablePay**: Verify "one amount per employee per type" rule via grouped sums
- **Rounding**: Spot-check that monthly salaries are whole numbers, other amounts are 2 decimals
- **EAO**: Verify EAO adjustments match manual calculations from 额外表 parameters
- **Diff columns**: All should be TRUE (or blank if not applicable); FALSE indicates mismatch

## Important Notes for Future Development

1. **Maintain late-bound COM**: Never add references to external libraries; keep all object types as `Object`
2. **Never bypass rounding services**: All numeric calculations must go through `RoundMonthlySalary()` or `RoundAmount2()`
3. **Log extensively**: Add `LogInfo` calls at major steps, `LogError` in all error handlers
4. **Test cross-month scenarios**: Many bugs occur at month boundaries (e.g., leaves spanning Feb-Mar)
5. **Update config before hard-coding**: If a new rule seems hard-coded, consider if it should be in config instead
6. **Export modules after changes**: Git only tracks `.bas` files, not the `.xlsm` workbook
7. **Chinese file names**: Use logical names in code; actual file names with Chinese characters are in config

## Documentation

- **User Guide (Chinese)**: `doc/HK_Payroll_Automation_操作说明.md`
- **Configuration Guide**: `doc/Configuration_Guide.md`
- **Quick Reference**: `doc/Quick_Reference.md`
- **Technical Implementation**: `guides/md/` folder
  - Subprocess 1 plan
  - Subprocess 2 plan
  - VBA guide (architecture and patterns)

## Contact & Support

This is an internal payroll automation tool. For questions about:
- Business rules: Contact HR/Payroll team
- Technical issues: Check `guides/md/FIXES_APPLIED.md` for recent fixes
- System errors: Review logs in `LogFolder` and `Log` sheet in control workbook
