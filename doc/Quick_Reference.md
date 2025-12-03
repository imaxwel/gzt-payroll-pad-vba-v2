# HK Payroll Automation - Quick Reference

## 宏命令速查

| 宏名称 | 功能 | 快捷键 |
|--------|------|--------|
| `Run_Subprocess1` | 运行子流程1，生成Flexi form output | Alt+F8 |
| `Run_Subprocess2` | 运行子流程2，生成Validation output | Alt+F8 |
| `Run_Both` | 依次运行两个子流程 | Alt+F8 |
| `SetupAll` | 初始化设置（首次使用） | Alt+F8 |
| `ValidateSetup` | 验证配置是否正确 | Alt+F8 |
| `TestConfiguration` | 测试配置加载 | Alt+F8 |

## 文件命名规范

### 输出文件
- Subprocess 1: `Flexi form out put YYYYMMDD.xlsx`
- Subprocess 2: `HK Payroll Validation Output YYYYMMDD.xlsx`

### 输入文件（必须）
```
1263 ADP flexiform template_HK_NewHire.xlsx
1263 ADP flexiform template_HK_Termination.xlsx
1263 ADP flexiform template_HK_DataChange.xlsx
1263 ADP flexiform template_HK_Comp.xlsx
1263 ADP flexiform template_HK_Attendance.xlsx
1263 ADP flexiform template_HK_Variable.xlsx
Employee_Leave_Transactions_Report.xlsx
EAO Summary Report_YYYYMM.xlsx
Workforce Detail - Payroll-AP.xlsx
额外表.xlsx
```

## 特殊月份标记

| 月份 | 标记 | 处理内容 |
|------|------|----------|
| 2月 | IsFlexBenefitMonth | 弹性福利 |
| 3月 | IsAIPMonth | 年度激励计划 |
| 5月 | IsRSUDivMonth | RSU全球股息 |
| 6月 | IsRSUEYMonth | RSU EY股息 |
| 8月 | IsFlexBenefitMonth | 弹性福利 |

## 舍入规则

| 类型 | 规则 |
|------|------|
| 月薪 | 四舍五入到整数 |
| 计算结果 | 四舍五入到2位小数 |
| Inspire Gross-up | 向上取整到整数 |

## 请假处理规则

| 请假类型 | 计算方式 | 特殊规则 |
|----------|----------|----------|
| 年假 | 工作日 | 排除周末和公众假期 |
| 病假 | 日历日 | 需≥4个连续工作日 |
| 无薪假 | 日历日 | - |
| PPTO | 日历日 | 当月结算上月 |
| 产假 | 日历日 | 需40周服务期 |
| 陪产假 | 日历日 | - |

## Diff列判断规则

| 情况 | 结果 |
|------|------|
| 两者都为空 | TRUE |
| Last Hire Date < 2025-01-01 | TRUE |
| 数值差异 < 0.01 | TRUE |
| 文本相同（忽略大小写） | TRUE |
| 其他不匹配 | FALSE |

## HC Check 公式

```
计算HC = 上月活跃HC - 上月离职(Included) - 本月离职(OC) + 本月新入职
```

## 常见错误代码

| 错误 | 原因 | 解决方案 |
|------|------|----------|
| File not found | 输入文件缺失 | 检查输入文件夹 |
| Column not found | 列名不匹配 | 检查输入文件列名 |
| Context not initialized | 未初始化 | 运行SetupAll |
| Invalid date | 日期格式错误 | 使用YYYY-MM-DD格式 |

## 日志级别

| 级别 | 颜色 | 说明 |
|------|------|------|
| INFO | 绿色 | 正常信息 |
| WARNING | 黄色 | 警告信息 |
| ERROR | 红色 | 错误信息 |

## 联系方式

如有问题，请联系IT支持团队。
