# modPathService 路径服务使用指南

## 概述

`modPathService` 模块提供分层的目录路径服务，支持按"年 + 期间类型（Month / Quarter / Adhoc）"组织输入文件，并支持当月/上月文件切换。

## 目录结构

```
...\Payroll_HK\
├── Input\
│   └── 2025\
│       ├── Month\
│       │   ├── 202501\
│       │   │   ├── Payroll Report.xlsx
│       │   │   ├── Workforce Detail - Payroll-AP.xlsx
│       │   │   ├── 1263 ADP flexiform template_HK_NewHire.xlsx
│       │   │   ├── 1263 ADP flexiform template_HK_Termination.xlsx
│       │   │   ├── Employee_Leave_Transactions_Report.xlsx
│       │   │   └── ...
│       │   ├── 202502\
│       │   │   └── ...
│       │   └── ...
│       ├── Quarter\
│       │   ├── 2025QX Payout Summary.xlsx
│       │   └── ...
│       └── Adhoc\
│           ├── Optional medical plan enrollment form.xlsx
│           └── Special_Bonus_List_2025.xlsx
└── Output\
    └── 2025\
        ├── HK_Payroll_Validation_Output_20250228.xlsx
        └── HK_Payroll_Validation_Output_20250329.xlsx
```

## 分层架构

| 层级 | 功能 | 主要函数 |
|------|------|----------|
| 第1层 | 基础路径初始化 | `InitPathService`, `InitPathServiceFromContext` |
| 第2层 | 期间计算（跨月/跨年） | `GetPeriodInfo`, `GetCurrentPeriodInfo`, `GetPreviousPeriodInfo` |
| 第3层 | 目录路径构建 | `BuildMonthlyInputPath`, `BuildQuarterlyInputPath`, `BuildAdhocInputPath` |
| 第4层 | 文件名映射 | `GetPhysicalFileName`, `GetFilePeriodType` |
| 第5层 | 统一文件路径接口 | `GetInputFilePathEx`, `GetCurrentMonthFilePath`, `GetPreviousMonthFilePath` |
| 第6层 | 兼容层 | `GetInputFilePathAuto`, `GetInputFilePathLegacy` |


## 使用示例

### 1. 基本用法 - 读取当月/上月文件

```vba
' 读取当月 Payroll Report
Dim currentPath As String
currentPath = GetInputFilePathEx("PayrollReport", poCurrentMonth)
' 结果: ...\Input\2025\Month\202502\Payroll Report.xlsx

' 读取上月 Payroll Report
Dim previousPath As String
previousPath = GetInputFilePathEx("PayrollReport", poPreviousMonth)
' 结果: ...\Input\2025\Month\202501\Payroll Report.xlsx

' 读取上月 Termination 文件
Dim termPath As String
termPath = GetInputFilePathEx("Termination", poPreviousMonth)
' 结果: ...\Input\2025\Month\202501\1263 ADP flexiform template_HK_Termination.xlsx
```

### 2. 便捷方法

```vba
' 当月文件
Dim path1 As String
path1 = GetCurrentMonthFilePath("NewHire")

' 上月文件
Dim path2 As String
path2 = GetPreviousMonthFilePath("NewHire")
```

### 3. 自动兼容旧结构

```vba
' 优先使用新结构，如果文件不存在则回退到旧结构
Dim autoPath As String
autoPath = GetInputFilePathAuto("PayrollReport", poCurrentMonth)
```

### 4. 在 HC Check 中使用

```vba
' 计算当月 Payroll HC
CalculatePayrollHC ws, poCurrentMonth

' 计算上月 Payroll HC
CalculatePayrollHC ws, poPreviousMonth
```

## 期间偏移量枚举

| 枚举值 | 说明 |
|--------|------|
| `poCurrentMonth` | 当月 (默认) |
| `poPreviousMonth` | 上月 |
| `poNextMonth` | 下月 (预留) |

## 期间类型枚举

| 枚举值 | 说明 | 目录 |
|--------|------|------|
| `ptMonth` | 月度文件 | `Year\Month\YYYYMM\` |
| `ptQuarter` | 季度文件 | `Year\Quarter\` |
| `ptAdhoc` | 临时文件 | `Year\Adhoc\` |

## 逻辑文件名映射

| 逻辑名称 | 物理文件名 | 期间类型 |
|----------|------------|----------|
| PayrollReport | Payroll Report.xlsx | Month |
| WorkforceDetail | Workforce Detail - Payroll-AP.xlsx | Month |
| NewHire | 1263 ADP flexiform template_HK_NewHire.xlsx | Month |
| Termination | 1263 ADP flexiform template_HK_Termination.xlsx | Month |
| DataChange | 1263 ADP flexiform template_HK_DataChange.xlsx | Month |
| Comp | 1263 ADP flexiform template_HK_Comp.xlsx | Month |
| Attendance | 1263 ADP flexiform template_HK_Attendance.xlsx | Month |
| Variable | 1263 ADP flexiform template_HK_Variable.xlsx | Month |
| EmployeeLeave | Employee_Leave_Transactions_Report.xlsx | Month |
| OneTimePayment | One time payment report.xlsx | Month |
| InspireAwards | Inspire Awards payroll report.xlsx | Month |
| EAOSummary | EAO Summary Report_YYYYMM.xlsx | Month |
| MerckPayroll | Merck Payroll Summary Report——xxx.xlsx | Month |
| SIPQIP | SIP QIP.xlsx | Month |
| FlexClaim | MSD HK Flex_Claim_Summary_Report.xlsx | Month |
| RSUGlobal | RSU Dividend global report.xlsx | Month |
| RSUEY | RSU Dividend EY report.xlsx | Month |
| DividendEY | Dividend EY report.xlsx | Month |
| AIPPayouts | AIP Payouts Payroll Report.xlsx | Month |
| ExtraTable | 额外表.xlsx | Month |
| AllowancePlan | Allowance plan report.xlsx | Month |
| QXPayout | 2025QX Payout Summary.xlsx | Quarter |
| OptionalMedical | Optional medical plan enrollment form.xlsx | Adhoc |

## 跨年处理

路径服务自动处理跨年场景：

```vba
' 假设当前薪资月为 202501
Dim prevInfo As tPeriodInfo
prevInfo = GetPeriodInfo("202501", poPreviousMonth)
' prevInfo.Year = 2024
' prevInfo.Month = 12
' prevInfo.YearMonth = "202412"
```

## 输入文件验证

Subprocess 2 在执行前会验证所有必需的输入文件是否存在（包括当月和上月）。

### 必需文件列表

| 文件 | 当月 | 上月 | 用途 |
|------|------|------|------|
| PayrollReport | ✓ | ✓ | Check Result 基准数据 / HC Check |
| WorkforceDetail | ✓ | - | Master Data Check |
| Termination | ✓ | ✓ | HC Check 计算 |
| NewHire | ✓ | ✓ | HC Check 计算 |

### 验证失败处理

如果任何必需文件缺失：
1. 在日志中记录详细的错误信息（包括缺失文件路径）
2. 弹出对话框提示用户
3. 终止流程执行，不生成输出文件

### 验证函数

```vba
' 在 modSP2_Main 中
Private Function ValidateRequiredInputFiles() As Boolean
    ' 验证当月和上月的必需文件
    ' 返回 True 表示所有文件存在
    ' 返回 False 表示有文件缺失
End Function
```

### 日志示例

```
[ERROR] modSP2_Main.ValidateRequiredInputFiles: 当月 Payroll Report 文件不存在: C:\...\Input\2025\Month\202502\Payroll Report.xlsx
[ERROR] modSP2_Main.ValidateRequiredInputFiles: 上月 Termination 文件不存在: C:\...\Input\2025\Month\202501\1263 ADP flexiform template_HK_Termination.xlsx
[ERROR] modSP2_Main.ValidateRequiredInputFiles: 以下必需输入文件缺失:
  - [当月] Payroll Report: C:\...\Input\2025\Month\202502\Payroll Report.xlsx
  - [上月] Termination: C:\...\Input\2025\Month\202501\1263 ADP flexiform template_HK_Termination.xlsx
```

## 迁移指南

1. 将现有输入文件按新目录结构重新组织
2. 使用 `GetInputFilePathAuto` 实现平滑过渡
3. 完成迁移后，可直接使用 `GetInputFilePathEx`
4. 确保当月和上月的必需文件都已准备好
