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
│       │   │   ├── Payroll_Report.xlsx
│       │   │   ├── Workforce_Detail.xlsx
│       │   │   ├── 1263_ADP_HK_NewHire.xlsx
│       │   │   ├── 1263_ADP_HK_Termination.xlsx
│       │   │   ├── Employee_Leave_Transactions.xlsx
│       │   │   └── ...
│       │   ├── 202502\
│       │   │   └── ...
│       │   └── ...
│       ├── Quarter\
│       │   ├── QX_Payout_Summary_2025Q1.xlsx
│       │   └── QX_Payout_Summary_2025Q2.xlsx
│       └── Adhoc\
│           ├── Optional_Medical_Upgrade_2025_List.xlsx
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
' 结果: ...\Input\2025\Month\202502\Payroll_Report.xlsx

' 读取上月 Payroll Report
Dim previousPath As String
previousPath = GetInputFilePathEx("PayrollReport", poPreviousMonth)
' 结果: ...\Input\2025\Month\202501\Payroll_Report.xlsx

' 读取上月 Termination 文件
Dim termPath As String
termPath = GetInputFilePathEx("Termination", poPreviousMonth)
' 结果: ...\Input\2025\Month\202501\1263_ADP_HK_Termination.xlsx
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
| PayrollReport | Payroll_Report.xlsx | Month |
| NewHire | 1263_ADP_HK_NewHire.xlsx | Month |
| Termination | 1263_ADP_HK_Termination.xlsx | Month |
| ExtraTable | Extra_Table.xlsx | Month |
| QXPayout | QX_Payout_Summary_YYYYQX.xlsx | Quarter |
| OptionalMedical | Optional_Medical_Upgrade_YYYY_List.xlsx | Adhoc |

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

## 迁移指南

1. 将现有输入文件按新目录结构重新组织
2. 使用 `GetInputFilePathAuto` 实现平滑过渡
3. 完成迁移后，可直接使用 `GetInputFilePathEx`
