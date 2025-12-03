# HK Payroll Automation - Configuration Guide

## Overview

This guide explains how to configure the HK Payroll Automation system for different payroll months and scenarios.

## Configuration Files

### 1. config.xlsx

Location: `config/config.xlsx`

#### PayrollSchedule Sheet

This sheet defines the payroll calendar for each month.

| Column | Description | Format |
|--------|-------------|--------|
| PayrollMonth | Payroll month identifier | YYYYMM |
| CutoffDate | Data cutoff date | YYYY-MM-DD |
| PayDate | Payment date | YYYY-MM-DD |
| IsAIPMonth | Annual Incentive Plan month flag | TRUE/FALSE |
| IsRSUDivMonth | RSU Dividend month flag | TRUE/FALSE |
| IsRSUEYMonth | RSU EY Dividend month flag | TRUE/FALSE |
| IsFlexBenefitMonth | Flexible Benefits month flag | TRUE/FALSE |

**Special Month Flags:**
- `IsAIPMonth`: Set to TRUE for March (AIP payout month)
- `IsRSUDivMonth`: Set to TRUE for May (RSU Global dividend)
- `IsRSUEYMonth`: Set to TRUE for June (RSU EY dividend)
- `IsFlexBenefitMonth`: Set to TRUE for February and August

#### Calendar Sheet

This sheet defines Hong Kong public holidays.

| Column | Description | Format |
|--------|-------------|--------|
| Date | Holiday date | YYYY-MM-DD |
| IsHKHoliday | Holiday flag | TRUE |

**2025 Hong Kong Public Holidays:**
- January 1 - New Year's Day
- January 29-31 - Lunar New Year
- April 4 - Ching Ming Festival
- April 18 - Good Friday
- April 19 - Day after Good Friday
- April 21 - Easter Monday
- May 1 - Labour Day
- May 5 - Buddha's Birthday
- June 2 - Tuen Ng Festival
- July 1 - HKSAR Establishment Day
- September 22 - Day after Mid-Autumn Festival
- October 1 - National Day
- October 7 - Chung Yeung Festival
- December 25 - Christmas Day
- December 26 - Boxing Day

#### ExchangeRates Sheet

| Column | Description |
|--------|-------------|
| RateName | Rate identifier |
| RateValue | Exchange rate value |

**Standard Rates:**
- `RSU_Global`: USD to HKD rate for RSU Global dividends
- `RSU_EY`: USD to HKD rate for RSU EY dividends
- `DefaultFX`: Default exchange rate (usually 1.0)

### 2. 额外表.xlsx

Location: `input/额外表.xlsx`

#### [需要每月维护] Sheet

Monthly maintenance parameters by WEIN.

| Column | Description |
|--------|-------------|
| WEIN | Employee WEIN |
| PPTO EAO Rate input | PPTO EAO rate |

#### [MPF&ORSO] Sheet

MPF and ORSO contribution parameters by WEIN.

| Column | Description |
|--------|-------------|
| WEIN | Employee WEIN |
| MPF EE VC % | MPF Employee Voluntary Contribution % |
| MPF ER VC % | MPF Employer Voluntary Contribution % |
| ORSO % | ORSO contribution % |
| ORSO ER Adj | ORSO Employer Adjustment |
| Percent Of ORSO ER | ORSO Employer % |
| Percent Of ORSO EE | ORSO Employee % |

#### [特殊奖金] Sheet

Special bonuses and allowances by WEIN.

| Column | Description |
|--------|-------------|
| WEIN | Employee WEIN |
| Flexible benefits | Flexible benefits amount |
| Other Allowance | Other allowance amount |
| Other Bonus | Other bonus amount |
| Other Rewards | Other rewards amount |
| Goods & Services Differential | G&S differential |

#### [Final payment] Sheet

Final payment parameters for terminated employees.

| Column | Description |
|--------|-------------|
| WEIN | Employee WEIN |
| MSD_or_Statutory | Policy type (MSD/Statutory) |
| TerminationType | Termination type |
| PILIndicator | Payment in Lieu indicator |
| NoticeGivenDate | Notice given date |
| NoticePeriod | Notice period (days) |
| Gratuities | Gratuities amount |
| BackPay | Back pay amount |

#### [Previous Month Terminated HC] Sheet

Previous month terminated headcount for HC Check.

| Column | Description |
|--------|-------------|
| WEIN | Employee WEIN |
| Month | Termination month |

## Runtime Configuration

### Runtime Sheet in HK_Payroll_Automation.xlsm

| Named Range | Description | Example |
|-------------|-------------|---------|
| InputFolder | Input files folder path | C:\HK_Payroll\input\ |
| OutputFolder | Output files folder path | C:\HK_Payroll\output\ |
| ConfigFolder | Config files folder path | C:\HK_Payroll\config\ |
| PayrollMonth | Target payroll month | 202501 |
| RunDate | Run date | 2025-01-15 |
| LogFolder | Log files folder path | C:\HK_Payroll\log\ |

**Important Notes:**
- All folder paths must end with a backslash `\`
- PayrollMonth must be in YYYYMM format
- RunDate should be the actual run date

## Business Rules Configuration

### Rounding Rules

Configured in `modRoundingService.bas`:

1. **Monthly Salary**: Rounded to nearest whole number (integer)
2. **All calculation results**: Rounded to 2 decimal places
3. **Inspire Points Gross-up**: Rounded UP to nearest integer

### Leave Processing Rules

Configured in `modSP1_Attendance.bas`:

1. **Annual Leave**: Split by calendar month, count business days only
2. **Sick Leave**: Requires 4+ consecutive business days
3. **Unpaid Leave**: Split by calendar month, count calendar days
4. **PPTO**: Current month settles previous month
5. **Maternity Leave**: Requires 40 weeks service check

### Grouping Rules

Configured in `modAggregationService.bas`:

- All variable pay items are grouped by Employee ID and Type
- Only one amount per employee per plan type is written
- Amounts are summed before writing

## Validation Rules

### Diff Column Rules

Configured in `modSP2_CheckResult_Diff.bas`:

1. **Both blank**: TRUE (match)
2. **Last Hired Date before 2025-01-01**: TRUE (special rule)
3. **Numeric comparison**: Allow 0.01 tolerance
4. **Text comparison**: Case-insensitive, trimmed

### HC Check Formula

```
Calculated HC = Previous Month Active HC 
              - Previous Month Terminated (Included)
              - Current Month Terminated (OC)
              + Current Month New Hire
```

## Troubleshooting Configuration Issues

### Common Configuration Errors

1. **Missing PayrollMonth in PayrollSchedule**
   - Add the missing month to config.xlsx

2. **Incorrect date format**
   - Ensure all dates are in YYYY-MM-DD format

3. **Missing exchange rate**
   - Add the rate to ExchangeRates sheet

4. **Path not found**
   - Verify folder paths exist and end with `\`

### Validation Checklist

Before running:
- [ ] PayrollMonth exists in PayrollSchedule
- [ ] All required input files are present
- [ ] 额外表.xlsx is updated for current month
- [ ] Exchange rates are current
- [ ] Calendar has current year holidays
