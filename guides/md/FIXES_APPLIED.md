# 修复说明 - 2024年12月4日

## 已修复的问题

### 1. VBA 日期格式错误

**问题描述：**
- 代码中使用了大写的 `"YYYYMMDD"` 和 `"YYYYMM"` 作为日期格式
- VBA 的 `Format()` 函数要求使用小写的 `"yyyymmdd"` 和 `"yyyymm"`
- 错误的格式导致生成奇怪的文件名（如 `2AB73400`）和路径访问错误

**已修复的文件：**

1. **code/modSP1_Main.bas** (第 48 行)
   - 修改前：`Format(G.Payroll.payDate, "YYYYMMDD")`
   - 修改后：`Format(G.Payroll.payDate, "yyyymmdd")`

2. **code/modSP2_Main.bas** (第 68 行)
   - 修改前：`Format(G.RunParams.RunDate, "YYYYMMDD")`
   - 修改后：`Format(G.RunParams.RunDate, "yyyymmdd")`

3. **code/modCalendarService.bas** (3处)
   - 第 103 行：`Format(currentStart, "YYYYMM")` → `Format(currentStart, "yyyymm")`
   - 第 145 行：`Format(currentStart, "YYYYMM")` → `Format(currentStart, "yyyymm")`
   - 第 172 行：`Format(d, "YYYYMM")` → `Format(d, "yyyymm")`

### 2. 配置文件缺失 202512 记录

**问题描述：**
- 当 `config.xlsx` 的 `PayrollSchedule` 工作表中找不到 202512 的记录时
- 系统会使用默认值：月末日期（2025年12月31日）
- 导致文件名变成 `Flexi form out put 20251231.xlsx` 而不是预期的日期

**需要手动操作：**

请打开 `config/config.xlsx` 文件，在 `PayrollSchedule` 工作表中添加 202512 的记录：

| PayrollMonth | CutoffDate | PayDate | 其他列... |
|--------------|------------|---------|-----------|
| 202512       | 2025-12-25 | 2025-12-26 | ... |

**注意：**
- `CutoffDate` 和 `PayDate` 请根据实际的薪资发放日期填写
- 如果你希望文件名使用特定日期，请确保 `PayDate` 列填写正确
- 日期格式应该是 Excel 的日期格式（如 2025-12-26）

## 测试建议

修复完成后，建议进行以下测试：

1. 确保 `config/config.xlsx` 中已添加 202512 的记录
2. 重新运行 Subprocess 1
3. 检查生成的文件名是否正确（应该是 `Flexi form out put 20251226.xlsx` 或你设置的日期）
4. 确认不再出现 `2AB73400` 这样的奇怪路径错误

## 其他发现

在检查代码时，还发现以下使用了正确格式的地方（无需修改）：
- `modEntryPoints.bas`：使用 `"yyyy-mm-dd"` 格式（正确）
- `modFormattingService.bas`：使用 `"yyyy/mm/dd"` 格式（正确）
- `modSP2_Main.bas`：使用 `"yyyy-mm-dd"` 格式（正确）

这些地方使用的是小写格式，所以是正确的。
