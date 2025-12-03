# HK Payroll Automation 操作说明

## 概述

本文档说明如何在 Windows 环境下部署和运行 HK Payroll Input & Validation 自动化系统。该系统包含两个子流程：

- **Subprocess 1**: 生成 Flexi form output 文件
- **Subprocess 2**: 生成 HK Payroll Validation Output 文件

## 系统要求

### 软件要求
- Windows 10/11
- Microsoft Excel 2016 或更高版本（需启用宏）
- Power Automate Desktop (PAD) - 可选，用于自动化调度

### 文件结构
```
HK_Payroll_Automation/
├── code/                          # VBA 模块文件
│   ├── modAppContext.bas
│   ├── modConfigService.bas
│   ├── modRoundingService.bas
│   ├── modCalendarService.bas
│   ├── modAggregationService.bas
│   ├── modEmployeeMappingService.bas
│   ├── modLoggingService.bas
│   ├── modFormattingService.bas
│   ├── modEAOService.bas
│   ├── modEntryPoints.bas
│   ├── modSP1_Main.bas
│   ├── modSP1_Attendance.bas
│   ├── modSP1_VariablePay.bas
│   ├── modSP2_Main.bas
│   ├── modSP2_CheckResult_*.bas
│   └── modSP2_HCCheck.bas
├── config/                        # 配置文件
│   └── config.xlsx
├── input/                         # 输入文件目录
├── output/                        # 输出文件目录
└── log/                           # 日志文件目录
```

## 安装步骤

### 1. 创建主控工作簿

1. 打开 Excel，创建新工作簿
2. 保存为 `HK_Payroll_Automation.xlsm`（启用宏的工作簿）
3. 按 `Alt + F11` 打开 VBA 编辑器

### 2. 导入 VBA 模块

在 VBA 编辑器中：

1. 右键点击项目名称
2. 选择 "导入文件..."
3. 依次导入 `code/` 目录下的所有 `.bas` 文件

导入顺序建议：
1. modAppContext.bas
2. modLoggingService.bas
3. modRoundingService.bas
4. modConfigService.bas
5. modCalendarService.bas
6. modAggregationService.bas
7. modEmployeeMappingService.bas
8. modEAOService.bas
9. modFormattingService.bas
10. modEntryPoints.bas
11. modSP1_Main.bas
12. modSP1_Attendance.bas
13. modSP1_VariablePay.bas
14. modSP2_Main.bas
15. modSP2_CheckResult_*.bas (所有检查模块)
16. modSP2_HCCheck.bas

### 3. 创建 Runtime 工作表

在主控工作簿中创建名为 `Runtime` 的工作表，设置以下命名区域：

| 命名区域 | 单元格 | 说明 |
|---------|--------|------|
| InputFolder | B2 | 输入文件夹路径 |
| OutputFolder | B3 | 输出文件夹路径 |
| ConfigFolder | B4 | 配置文件夹路径 |
| PayrollMonth | B5 | 薪资月份 (YYYYMM) |
| RunDate | B6 | 运行日期 |
| LogFolder | B7 | 日志文件夹路径 |
| SP_Status | B8 | 运行状态 |
| SP_Message | B9 | 状态消息 |

示例配置：
```
A列          B列
InputFolder  C:\HK_Payroll\input\
OutputFolder C:\HK_Payroll\output\
ConfigFolder C:\HK_Payroll\config\
PayrollMonth 202501
RunDate      2025-01-15
LogFolder    C:\HK_Payroll\log\
```

### 4. 创建配置文件

创建 `config/config.xlsx`，包含以下工作表：

#### PayrollSchedule 工作表
| PayrollMonth | CutoffDate | PayDate | IsAIPMonth | IsRSUDivMonth | IsFlexBenefitMonth |
|--------------|------------|---------|------------|---------------|-------------------|
| 202501 | 2025-01-25 | 2025-01-31 | FALSE | FALSE | FALSE |
| 202502 | 2025-02-25 | 2025-02-28 | FALSE | FALSE | TRUE |
| 202503 | 2025-03-25 | 2025-03-31 | TRUE | FALSE | FALSE |

#### Calendar 工作表
| Date | IsHKHoliday |
|------|-------------|
| 2025-01-01 | TRUE |
| 2025-01-29 | TRUE |
| ... | ... |

#### ExchangeRates 工作表
| RateName | RateValue |
|----------|-----------|
| RSU_Global | 7.8 |
| RSU_EY | 7.8 |
| DefaultFX | 1.0 |

## 运行说明

### 手动运行

#### 运行 Subprocess 1
1. 打开 `HK_Payroll_Automation.xlsm`
2. 确认 Runtime 工作表中的参数正确
3. 按 `Alt + F8` 打开宏对话框
4. 选择 `Run_Subprocess1`
5. 点击 "运行"

#### 运行 Subprocess 2
1. 确保 Subprocess 1 已成功完成
2. 按 `Alt + F8` 打开宏对话框
3. 选择 `Run_Subprocess2`
4. 点击 "运行"

#### 运行两个子流程
1. 按 `Alt + F8` 打开宏对话框
2. 选择 `Run_Both`
3. 点击 "运行"

### 使用 Power Automate Desktop (PAD)

#### PAD 流程配置

1. 创建新的 PAD 流程
2. 添加以下步骤：

```
1. 设置变量
   - vInputFolder = "C:\HK_Payroll\input\"
   - vOutputFolder = "C:\HK_Payroll\output\"
   - vConfigFolder = "C:\HK_Payroll\config\"
   - vPayrollMonth = "202501"
   - vRunDate = 当前日期

2. 启动 Excel
   - 打开 HK_Payroll_Automation.xlsm

3. 写入运行参数
   - 将变量写入 Runtime 工作表

4. 运行 Excel 宏
   - 宏名称: Run_Subprocess1 或 Run_Subprocess2

5. 读取状态
   - 从 Runtime!SP_Status 读取结果

6. 条件判断
   - 如果状态 = "ERROR"，发送通知

7. 关闭 Excel
```

## 输入文件说明

### Subprocess 1 输入文件

| 文件名 | 说明 |
|--------|------|
| 1263 ADP flexiform template_HK_NewHire.xlsx | 新员工数据 |
| 1263 ADP flexiform template_HK_Termination.xlsx | 离职数据 |
| 1263 ADP flexiform template_HK_DataChange.xlsx | 信息变更 |
| 1263 ADP flexiform template_HK_Comp.xlsx | 薪资变更 |
| 1263 ADP flexiform template_HK_Attendance.xlsx | 考勤数据 |
| 1263 ADP flexiform template_HK_Variable.xlsx | 可变薪酬 |
| One time payment report.xlsx | 一次性支付 |
| Inspire Awards payroll report.xlsx | Inspire 奖励 |
| Employee_Leave_Transactions_Report.xlsx | 请假记录 |
| EAO Summary Report_YYYYMM.xlsx | EAO 汇总 |
| Workforce Detail - Payroll-AP.xlsx | 员工详情 |
| Merck Payroll Summary Report——xxx.xlsx | 薪资汇总 |
| SIP QIP.xlsx | 销售激励 |
| MSD HK Flex_Claim_Summary_Report.xlsx | 弹性福利 |
| RSU Dividend global report.xlsx | RSU 股息(全球) |
| Dividend EY report.xlsx | RSU 股息(EY) |
| AIP Payouts Payroll Report.xlsx | AIP 支付 |
| 额外表.xlsx | 额外参数表 |

### Subprocess 2 输入文件

除 Subprocess 1 的文件外，还需要：

| 文件名 | 说明 |
|--------|------|
| Payroll Report.xlsx | 当月薪资报告 |
| Allowance plan report.xlsx | 津贴计划 |
| 2025QX Payout Summary.xlsx | 季度支付汇总 |
| Optional medical plan enrollment form.xlsx | 医疗计划 |

## 输出文件说明

### Subprocess 1 输出
- `Flexi form out put YYYYMMDD.xlsx`
  - RunSummary: 运行摘要
  - NewHire: 新员工
  - InformationChange: 信息变更
  - SalaryChange: 薪资变更
  - Termination: 离职
  - Attendance: 考勤
  - VariablePay: 可变薪酬

### Subprocess 2 输出
- `HK Payroll Validation Output YYYYMMDD.xlsx`
  - Check Result: 验证结果（Benchmark/Check/Diff 列）
  - HC Check: 人数核对

## 故障排除

### 常见问题

#### 1. 宏被禁用
- 打开 Excel 选项 → 信任中心 → 信任中心设置
- 选择 "启用所有宏"

#### 2. 文件未找到错误
- 检查 Runtime 工作表中的路径是否正确
- 确保路径以反斜杠 `\` 结尾
- 确认所有必需的输入文件都存在

#### 3. 运行时错误
- 查看 Log 工作表中的错误信息
- 检查日志文件夹中的 .log 文件

#### 4. 数据不匹配
- 确认输入文件的列名与预期一致
- 检查日期格式是否正确

### 日志查看

运行日志保存在：
1. 主控工作簿的 `Log` 工作表
2. 日志文件夹中的 `.log` 文件

日志格式：
```
时间戳    级别    消息
2025-01-15 10:30:00    INFO    [modSP1_Main.SP1_Execute] Starting...
2025-01-15 10:30:05    ERROR   [modSP1_Main.CopyFlexiformSheet] File not found...
```

## 维护说明

### 更新配置
- 每月更新 `config.xlsx` 中的 PayrollSchedule
- 每年更新 Calendar 工作表中的香港公众假期
- 根据需要更新汇率

### 更新额外表
- 每月更新 `额外表.xlsx` 中的：
  - [需要每月维护] - PPTO EAO Rate
  - [MPF&ORSO] - MPF/ORSO 参数
  - [特殊奖金] - 特殊奖金项目
  - [Final payment] - 离职支付参数

## 联系支持

如有问题，请联系 IT 支持团队。
