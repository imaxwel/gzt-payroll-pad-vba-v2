Attribute VB_Name = "modPathService"
'==============================================================================
' Module: modPathService
' Purpose: 分层目录路径服务 - 支持按年/期间类型(Month/Quarter/Adhoc)组织输入文件
' Description: 提供统一的文件路径解析接口，支持当月/上月切换，
'              跨月、跨季度、跨年逻辑独立分层，便于后续目录结构调整
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Enum: ePeriodOffset
' Purpose: 期间偏移量枚举
'------------------------------------------------------------------------------
Public Enum ePeriodOffset
    poCurrentMonth = 0      ' 当月
    poPreviousMonth = -1    ' 上月
    poNextMonth = 1         ' 下月 (预留)
End Enum

'------------------------------------------------------------------------------
' Enum: ePeriodType
' Purpose: 期间类型枚举
'------------------------------------------------------------------------------
Public Enum ePeriodType
    ptMonth = 1             ' 月度文件
    ptQuarter = 2           ' 季度文件
    ptAdhoc = 3             ' 一次性/临时文件
End Enum

'------------------------------------------------------------------------------
' Type: tPeriodInfo
' Purpose: 期间信息结构
'------------------------------------------------------------------------------
Public Type tPeriodInfo
    Year As Integer         ' 年份 (e.g., 2025)
    Month As Integer        ' 月份 (1-12)
    Quarter As Integer      ' 季度 (1-4)
    YearMonth As String     ' "YYYYMM" 格式
    YearQuarter As String   ' "YYYYQX" 格式
End Type

' 模块级缓存 - 基础路径
Private mBaseInputPath As String
Private mBaseOutputPath As String
Private mIsPathInitialized As Boolean


'==============================================================================
' 第一层：基础路径初始化
'==============================================================================

'------------------------------------------------------------------------------
' Sub: InitPathService
' Purpose: 初始化路径服务，设置基础路径
' Parameters:
'   baseInputPath - 输入文件根目录 (e.g., "C:\Payroll_HK\Input\")
'   baseOutputPath - 输出文件根目录 (e.g., "C:\Payroll_HK\Output\")
'------------------------------------------------------------------------------
Public Sub InitPathService(baseInputPath As String, baseOutputPath As String)
    mBaseInputPath = EnsureTrailingSlash(baseInputPath)
    mBaseOutputPath = EnsureTrailingSlash(baseOutputPath)
    mIsPathInitialized = True
    
    LogInfo "modPathService", "InitPathService", _
        "Path service initialized. Input: " & mBaseInputPath & ", Output: " & mBaseOutputPath
End Sub

'------------------------------------------------------------------------------
' Sub: InitPathServiceFromContext
' Purpose: 从全局上下文初始化路径服务
'------------------------------------------------------------------------------
Public Sub InitPathServiceFromContext()
    If G.IsInitialised Then
        ' 假设 RunParams.InputFolder 是新结构的根目录
        ' 如果是旧结构，需要调整
        InitPathService G.RunParams.InputFolder, G.RunParams.OutputFolder
    Else
        Err.Raise vbObjectError + 1001, "InitPathServiceFromContext", _
            "Global context not initialized"
    End If
End Sub

'==============================================================================
' 第二层：期间计算逻辑 (跨月/跨季度/跨年)
'==============================================================================

'------------------------------------------------------------------------------
' Function: GetPeriodInfo
' Purpose: 根据基准期间和偏移量计算目标期间信息
' Parameters:
'   baseYearMonth - 基准年月 "YYYYMM"
'   offset - 期间偏移量 (ePeriodOffset)
' Returns: tPeriodInfo 结构
'------------------------------------------------------------------------------
Public Function GetPeriodInfo(baseYearMonth As String, offset As ePeriodOffset) As tPeriodInfo
    Dim info As tPeriodInfo
    Dim baseYear As Integer, baseMonth As Integer
    Dim targetDate As Date
    
    ' 解析基准年月
    baseYear = CInt(Left(baseYearMonth, 4))
    baseMonth = CInt(Right(baseYearMonth, 2))
    
    ' 计算目标日期 (使用 DateAdd 自动处理跨年)
    targetDate = DateAdd("m", CLng(offset), DateSerial(baseYear, baseMonth, 1))
    
    ' 填充期间信息
    info.Year = Year(targetDate)
    info.Month = Month(targetDate)
    info.Quarter = GetQuarterFromMonth(info.Month)
    info.YearMonth = Format(targetDate, "YYYYMM")
    info.YearQuarter = CStr(info.Year) & "Q" & CStr(info.Quarter)
    
    GetPeriodInfo = info
End Function


'------------------------------------------------------------------------------
' Function: GetQuarterFromMonth
' Purpose: 根据月份获取季度
'------------------------------------------------------------------------------
Private Function GetQuarterFromMonth(mo As Integer) As Integer
    GetQuarterFromMonth = ((mo - 1) \ 3) + 1
End Function

'------------------------------------------------------------------------------
' Function: GetCurrentPeriodInfo
' Purpose: 获取当前薪资月的期间信息
'------------------------------------------------------------------------------
Public Function GetCurrentPeriodInfo() As tPeriodInfo
    GetCurrentPeriodInfo = GetPeriodInfo(G.Payroll.payrollMonth, poCurrentMonth)
End Function

'------------------------------------------------------------------------------
' Function: GetPreviousPeriodInfo
' Purpose: 获取上月的期间信息
'------------------------------------------------------------------------------
Public Function GetPreviousPeriodInfo() As tPeriodInfo
    GetPreviousPeriodInfo = GetPeriodInfo(G.Payroll.payrollMonth, poPreviousMonth)
End Function

'==============================================================================
' 第三层：目录路径构建
'==============================================================================

'------------------------------------------------------------------------------
' Function: BuildMonthlyInputPath
' Purpose: 构建月度输入文件目录路径
' Parameters:
'   periodInfo - 期间信息
' Returns: 完整目录路径 (e.g., "...\2025\Month\202501\")
'------------------------------------------------------------------------------
Public Function BuildMonthlyInputPath(periodInfo As tPeriodInfo) As String
    Dim path As String
    
    path = mBaseInputPath
    path = path & CStr(periodInfo.Year) & "\"
    path = path & "Month\"
    path = path & periodInfo.YearMonth & "\"
    
    BuildMonthlyInputPath = path
End Function

'------------------------------------------------------------------------------
' Function: BuildQuarterlyInputPath
' Purpose: 构建季度输入文件目录路径
' Parameters:
'   periodInfo - 期间信息
' Returns: 完整目录路径 (e.g., "...\2025\Quarter\")
'------------------------------------------------------------------------------
Public Function BuildQuarterlyInputPath(periodInfo As tPeriodInfo) As String
    Dim path As String
    
    path = mBaseInputPath
    path = path & CStr(periodInfo.Year) & "\"
    path = path & "Quarter\"
    
    BuildQuarterlyInputPath = path
End Function

'------------------------------------------------------------------------------
' Function: BuildAdhocInputPath
' Purpose: 构建临时/一次性输入文件目录路径
' Parameters:
'   periodInfo - 期间信息
' Returns: 完整目录路径 (e.g., "...\2025\Adhoc\")
'------------------------------------------------------------------------------
Public Function BuildAdhocInputPath(periodInfo As tPeriodInfo) As String
    Dim path As String
    
    path = mBaseInputPath
    path = path & CStr(periodInfo.Year) & "\"
    path = path & "Adhoc\"
    
    BuildAdhocInputPath = path
End Function


'------------------------------------------------------------------------------
' Function: BuildOutputPath
' Purpose: 构建输出文件目录路径
' Parameters:
'   periodInfo - 期间信息
' Returns: 完整目录路径 (e.g., "...\Output\2025\")
'------------------------------------------------------------------------------
Public Function BuildOutputPath(periodInfo As tPeriodInfo) As String
    Dim path As String
    
    path = mBaseOutputPath
    path = path & CStr(periodInfo.Year) & "\"
    
    BuildOutputPath = path
End Function

'==============================================================================
' 第四层：文件名映射 (逻辑名称 -> 物理文件名)
'==============================================================================

'------------------------------------------------------------------------------
' Function: GetPhysicalFileName
' Purpose: 根据逻辑名称获取物理文件名
' Parameters:
'   logicalName - 逻辑文件名
'   periodInfo - 期间信息 (用于动态文件名)
' Returns: 物理文件名
'------------------------------------------------------------------------------
Public Function GetPhysicalFileName(logicalName As String, periodInfo As tPeriodInfo) As String
    Dim fileName As String
    
    Select Case UCase(logicalName)
        ' === 月度文件 (Monthly) ===
        Case "PAYROLLREPORT"
            fileName = "Payroll_Report.xlsx"
        Case "WORKFORCEDETAIL"
            fileName = "Workforce_Detail.xlsx"
        Case "NEWHIRE"
            fileName = "1263_ADP_HK_NewHire.xlsx"
        Case "TERMINATION"
            fileName = "1263_ADP_HK_Termination.xlsx"
        Case "DATACHANGE"
            fileName = "1263_ADP_HK_DataChange.xlsx"
        Case "COMP"
            fileName = "1263_ADP_HK_Comp.xlsx"
        Case "ATTENDANCE"
            fileName = "1263_ADP_HK_Attendance.xlsx"
        Case "VARIABLE"
            fileName = "1263_ADP_HK_Variable.xlsx"
        Case "EMPLOYEELEAVE"
            fileName = "Employee_Leave_Transactions.xlsx"
        Case "ONETIMEPAYMENT"
            fileName = "One_Time_Payment.xlsx"
        Case "INSPIREWARDS"
            fileName = "Inspire_Awards.xlsx"
        Case "EAOSUMMARY"
            fileName = "EAO_Summary_Report.xlsx"
        Case "MERCKPAYROLL"
            fileName = "Merck_Payroll_Summary.xlsx"
        Case "SIPQIP"
            fileName = "SIP_QIP.xlsx"
        Case "FLEXCLAIM"
            fileName = "Flex_Claim_Summary.xlsx"
        Case "RSUGLOBAL"
            fileName = "RSU_Dividend_Global.xlsx"
        Case "RSUEY"
            fileName = "RSU_Dividend_EY.xlsx"
        Case "AIPPAYOUTS"
            fileName = "AIP_Payouts.xlsx"
        Case "EXTRATABLE"
            fileName = "Extra_Table.xlsx"
        Case "ALLOWANCEPLAN"
            fileName = "Allowance_Plan.xlsx"
            
        ' === 季度文件 (Quarterly) - 文件名包含季度信息 ===
        Case "QXPAYOUT"
            fileName = "QX_Payout_Summary_" & periodInfo.YearQuarter & ".xlsx"
            
        ' === 临时文件 (Adhoc) ===
        Case "OPTIONALMEDICAL"
            fileName = "Optional_Medical_Upgrade_" & CStr(periodInfo.Year) & "_List.xlsx"
        Case "SPECIALBONUS"
            fileName = "Special_Bonus_List_" & CStr(periodInfo.Year) & ".xlsx"
            
        Case Else
            fileName = logicalName
    End Select
    
    GetPhysicalFileName = fileName
End Function


'------------------------------------------------------------------------------
' Function: GetFilePeriodType
' Purpose: 根据逻辑名称获取文件的期间类型
' Parameters:
'   logicalName - 逻辑文件名
' Returns: ePeriodType
'------------------------------------------------------------------------------
Public Function GetFilePeriodType(logicalName As String) As ePeriodType
    Select Case UCase(logicalName)
        ' 季度文件
        Case "QXPAYOUT"
            GetFilePeriodType = ptQuarter
        ' 临时文件
        Case "OPTIONALMEDICAL", "SPECIALBONUS"
            GetFilePeriodType = ptAdhoc
        ' 默认为月度文件
        Case Else
            GetFilePeriodType = ptMonth
    End Select
End Function

'==============================================================================
' 第五层：统一文件路径接口 (对外暴露的主要API)
'==============================================================================

'------------------------------------------------------------------------------
' Function: GetInputFilePathEx
' Purpose: 获取输入文件完整路径 (支持期间偏移)
' Parameters:
'   logicalName - 逻辑文件名
'   offset - 期间偏移量 (默认当月)
' Returns: 完整文件路径
' Example:
'   GetInputFilePathEx("PayrollReport", poCurrentMonth)  -> 当月Payroll Report
'   GetInputFilePathEx("PayrollReport", poPreviousMonth) -> 上月Payroll Report
'   GetInputFilePathEx("Termination", poPreviousMonth)   -> 上月Termination
'------------------------------------------------------------------------------
Public Function GetInputFilePathEx(logicalName As String, _
                                   Optional offset As ePeriodOffset = poCurrentMonth) As String
    Dim periodInfo As tPeriodInfo
    Dim basePath As String
    Dim fileName As String
    Dim periodType As ePeriodType
    
    ' 确保路径服务已初始化
    EnsurePathInitialized
    
    ' 获取目标期间信息
    periodInfo = GetPeriodInfo(G.Payroll.payrollMonth, offset)
    
    ' 获取文件期间类型
    periodType = GetFilePeriodType(logicalName)
    
    ' 根据期间类型构建基础路径
    Select Case periodType
        Case ptMonth
            basePath = BuildMonthlyInputPath(periodInfo)
        Case ptQuarter
            basePath = BuildQuarterlyInputPath(periodInfo)
        Case ptAdhoc
            basePath = BuildAdhocInputPath(periodInfo)
    End Select
    
    ' 获取物理文件名
    fileName = GetPhysicalFileName(logicalName, periodInfo)
    
    GetInputFilePathEx = basePath & fileName
End Function

'------------------------------------------------------------------------------
' Function: GetCurrentMonthFilePath
' Purpose: 获取当月输入文件路径 (便捷方法)
'------------------------------------------------------------------------------
Public Function GetCurrentMonthFilePath(logicalName As String) As String
    GetCurrentMonthFilePath = GetInputFilePathEx(logicalName, poCurrentMonth)
End Function

'------------------------------------------------------------------------------
' Function: GetPreviousMonthFilePath
' Purpose: 获取上月输入文件路径 (便捷方法)
'------------------------------------------------------------------------------
Public Function GetPreviousMonthFilePath(logicalName As String) As String
    GetPreviousMonthFilePath = GetInputFilePathEx(logicalName, poPreviousMonth)
End Function


'------------------------------------------------------------------------------
' Function: GetOutputFilePath
' Purpose: 获取输出文件完整路径
' Parameters:
'   fileName - 输出文件名
' Returns: 完整文件路径
'------------------------------------------------------------------------------
Public Function GetOutputFilePath(fileName As String) As String
    Dim periodInfo As tPeriodInfo
    
    EnsurePathInitialized
    
    periodInfo = GetCurrentPeriodInfo()
    GetOutputFilePath = BuildOutputPath(periodInfo) & fileName
End Function

'==============================================================================
' 第六层：兼容层 - 支持旧目录结构 (可选)
'==============================================================================

'------------------------------------------------------------------------------
' Function: GetInputFilePathLegacy
' Purpose: 兼容旧目录结构的文件路径获取
' Note: 当目录结构未迁移时使用此方法
'------------------------------------------------------------------------------
Public Function GetInputFilePathLegacy(logicalName As String) As String
    ' 直接调用原有的 GetInputFilePath 函数
    GetInputFilePathLegacy = GetInputFilePath(logicalName)
End Function

'------------------------------------------------------------------------------
' Function: GetInputFilePathAuto
' Purpose: 自动检测目录结构并返回正确路径
' Parameters:
'   logicalName - 逻辑文件名
'   offset - 期间偏移量
' Returns: 完整文件路径 (优先新结构，回退旧结构)
'------------------------------------------------------------------------------
Public Function GetInputFilePathAuto(logicalName As String, _
                                     Optional offset As ePeriodOffset = poCurrentMonth) As String
    Dim newPath As String
    Dim legacyPath As String
    
    ' 尝试新结构路径
    On Error Resume Next
    newPath = GetInputFilePathEx(logicalName, offset)
    On Error GoTo 0
    
    ' 检查新路径文件是否存在
    If Len(newPath) > 0 And Dir(newPath) <> "" Then
        GetInputFilePathAuto = newPath
        Exit Function
    End If
    
    ' 回退到旧结构 (仅当月有效)
    If offset = poCurrentMonth Then
        legacyPath = GetInputFilePathLegacy(logicalName)
        If Dir(legacyPath) <> "" Then
            GetInputFilePathAuto = legacyPath
            Exit Function
        End If
    End If
    
    ' 返回新结构路径 (即使文件不存在，让调用方处理)
    GetInputFilePathAuto = newPath
End Function

'==============================================================================
' 辅助函数
'==============================================================================

'------------------------------------------------------------------------------
' Sub: EnsurePathInitialized
' Purpose: 确保路径服务已初始化
'------------------------------------------------------------------------------
Private Sub EnsurePathInitialized()
    If Not mIsPathInitialized Then
        ' 尝试从全局上下文初始化
        If G.IsInitialised Then
            InitPathServiceFromContext
        Else
            Err.Raise vbObjectError + 1002, "EnsurePathInitialized", _
                "Path service not initialized. Call InitPathService first."
        End If
    End If
End Sub

'------------------------------------------------------------------------------
' Function: EnsureTrailingSlash
' Purpose: 确保路径以反斜杠结尾
'------------------------------------------------------------------------------
Private Function EnsureTrailingSlash(path As String) As String
    If Len(path) > 0 And Right(path, 1) <> "\" Then
        EnsureTrailingSlash = path & "\"
    Else
        EnsureTrailingSlash = path
    End If
End Function


'------------------------------------------------------------------------------
' Sub: EnsureInputFolderExists
' Purpose: 确保输入目录存在
'------------------------------------------------------------------------------
Public Sub EnsureInputFolderExists(logicalName As String, _
                                   Optional offset As ePeriodOffset = poCurrentMonth)
    Dim periodInfo As tPeriodInfo
    Dim folderPath As String
    Dim periodType As ePeriodType
    
    EnsurePathInitialized
    
    periodInfo = GetPeriodInfo(G.Payroll.payrollMonth, offset)
    periodType = GetFilePeriodType(logicalName)
    
    Select Case periodType
        Case ptMonth
            folderPath = BuildMonthlyInputPath(periodInfo)
        Case ptQuarter
            folderPath = BuildQuarterlyInputPath(periodInfo)
        Case ptAdhoc
            folderPath = BuildAdhocInputPath(periodInfo)
    End Select
    
    EnsureFolderExists folderPath
End Sub

'------------------------------------------------------------------------------
' Sub: EnsureOutputFolderExists
' Purpose: 确保输出目录存在
'------------------------------------------------------------------------------
Public Sub EnsureOutputFolderExists()
    Dim periodInfo As tPeriodInfo
    Dim folderPath As String
    
    EnsurePathInitialized
    
    periodInfo = GetCurrentPeriodInfo()
    folderPath = BuildOutputPath(periodInfo)
    
    EnsureFolderExists folderPath
End Sub

'------------------------------------------------------------------------------
' Function: FileExistsForPeriod
' Purpose: 检查指定期间的文件是否存在
'------------------------------------------------------------------------------
Public Function FileExistsForPeriod(logicalName As String, _
                                    Optional offset As ePeriodOffset = poCurrentMonth) As Boolean
    Dim filePath As String
    
    filePath = GetInputFilePathEx(logicalName, offset)
    FileExistsForPeriod = (Dir(filePath) <> "")
End Function

'------------------------------------------------------------------------------
' Function: GetPeriodDescription
' Purpose: 获取期间描述文本 (用于日志和UI)
'------------------------------------------------------------------------------
Public Function GetPeriodDescription(offset As ePeriodOffset) As String
    Select Case offset
        Case poCurrentMonth
            GetPeriodDescription = "当月"
        Case poPreviousMonth
            GetPeriodDescription = "上月"
        Case poNextMonth
            GetPeriodDescription = "下月"
        Case Else
            GetPeriodDescription = "未知期间"
    End Select
End Function

'------------------------------------------------------------------------------
' Sub: LogPathInfo
' Purpose: 记录路径信息到日志 (调试用)
'------------------------------------------------------------------------------
Public Sub LogPathInfo(logicalName As String, offset As ePeriodOffset)
    Dim filePath As String
    Dim periodDesc As String
    Dim exists As String
    
    filePath = GetInputFilePathEx(logicalName, offset)
    periodDesc = GetPeriodDescription(offset)
    exists = IIf(Dir(filePath) <> "", "存在", "不存在")
    
    LogInfo "modPathService", "LogPathInfo", _
        logicalName & " (" & periodDesc & "): " & filePath & " [" & exists & "]"
End Sub

