## HK Payroll Form PAD Workflow

This workflow launches the control workbook and displays the VBA payroll form.

### Prerequisites
- Excel installed on the machine running PAD.
- `HK_Payroll_Automation.xlsm` contains all VBA modules from `code/`.
- Macro security allows running signed/trusted macros.

### Steps (PAD Designer)
1. **Display select file dialog**
   - Action: `Display > Select file dialog`
   - Title: `Select HK_Payroll_Automation.xlsm`
   - Filter: `Excel Macro-Enabled Workbook (*.xlsm)`
   - Store selected path into variable `payrollWbPath`
   - Store button pressed into variable `btnPressed`
2. **Condition**
   - If `btnPressed` is not `Open`, then:
     - `Display message dialog` with text `No workbook selected. Workflow ends.`
     - `Exit` with code 0
3. **Launch Excel**
   - Action: `Excel > Launch Excel`
   - Launch and open under existing process: **Enabled**
   - Path: `%payrollWbPath%`
   - Visible: `True`
   - ReadOnly: `False`
   - Store instance into variable `PayrollExcel`
4. **Run Excel macro**
   - Action: `Excel > Run Excel macro`
   - Instance: `%PayrollExcel%`
   - Macro name: `startformMain` (alias of `ShowPayrollForm`)
5. **End**
   - Leave Excel open for user interaction in the form.

### Notes
- The VBA form handles refreshing file paths, validation, and running subprocesses.
- PAD does not need to set Runtime parameters before showing the form.

