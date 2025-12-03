Attribute VB_Name = "modRoundingService"
'==============================================================================
' Module: modRoundingService
' Purpose: Centralized rounding functions for all calculations
' Description: Ensures consistent rounding rules across Subprocess 1 and 2
'              - Monthly Salary: rounded to whole number (integer)
'              - All calculation results: rounded to 2 decimal places
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Function: RoundMonthlySalary
' Purpose: Round monthly salary to nearest whole number
' Parameters:
'   v - Value to round (Variant to handle various input types)
' Returns: Rounded value as Double (integer value)
' Note: This should be applied ONCE when salary is first read, then used
'       consistently in all subsequent calculations
'------------------------------------------------------------------------------
Public Function RoundMonthlySalary(ByVal v As Variant) As Double
    On Error Resume Next
    If IsNumeric(v) And Not IsEmpty(v) Then
        RoundMonthlySalary = WorksheetFunction.Round(CDbl(v), 0)
    Else
        RoundMonthlySalary = 0
    End If
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Function: RoundAmount2
' Purpose: Round amount to 2 decimal places
' Parameters:
'   v - Value to round (Variant to handle various input types)
' Returns: Rounded value as Double
' Note: Use this for all calculation results (pay items, adjustments, etc.)
'------------------------------------------------------------------------------
Public Function RoundAmount2(ByVal v As Variant) As Double
    On Error Resume Next
    If IsNumeric(v) And Not IsEmpty(v) Then
        RoundAmount2 = WorksheetFunction.Round(CDbl(v), 2)
    Else
        RoundAmount2 = 0
    End If
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Function: SafeAdd2
' Purpose: Safely add two values and round to 2 decimal places
' Parameters:
'   a, b - Values to add (handles Null/Empty)
' Returns: Sum rounded to 2 decimal places
'------------------------------------------------------------------------------
Public Function SafeAdd2(ByVal a As Variant, ByVal b As Variant) As Double
    SafeAdd2 = RoundAmount2(Nz(a, 0) + Nz(b, 0))
End Function

'------------------------------------------------------------------------------
' Function: SafeSubtract2
' Purpose: Safely subtract two values and round to 2 decimal places
' Parameters:
'   a, b - Values (a - b), handles Null/Empty
' Returns: Difference rounded to 2 decimal places
'------------------------------------------------------------------------------
Public Function SafeSubtract2(ByVal a As Variant, ByVal b As Variant) As Double
    SafeSubtract2 = RoundAmount2(Nz(a, 0) - Nz(b, 0))
End Function

'------------------------------------------------------------------------------
' Function: SafeMultiply2
' Purpose: Safely multiply two values and round to 2 decimal places
' Parameters:
'   a, b - Values to multiply (handles Null/Empty)
' Returns: Product rounded to 2 decimal places
'------------------------------------------------------------------------------
Public Function SafeMultiply2(ByVal a As Variant, ByVal b As Variant) As Double
    SafeMultiply2 = RoundAmount2(Nz(a, 0) * Nz(b, 0))
End Function

'------------------------------------------------------------------------------
' Function: SafeDivide2
' Purpose: Safely divide two values and round to 2 decimal places
' Parameters:
'   a - Numerator
'   b - Denominator (returns 0 if zero or empty)
' Returns: Quotient rounded to 2 decimal places
'------------------------------------------------------------------------------
Public Function SafeDivide2(ByVal a As Variant, ByVal b As Variant) As Double
    Dim denominator As Double
    denominator = Nz(b, 0)
    
    If denominator = 0 Then
        SafeDivide2 = 0
    Else
        SafeDivide2 = RoundAmount2(Nz(a, 0) / denominator)
    End If
End Function

'------------------------------------------------------------------------------
' Function: RoundUpInteger
' Purpose: Round up to nearest integer (for specific cases like Inspire gross-up)
' Parameters:
'   v - Value to round up
' Returns: Rounded up integer value as Double
'------------------------------------------------------------------------------
Public Function RoundUpInteger(ByVal v As Variant) As Double
    On Error Resume Next
    If IsNumeric(v) And Not IsEmpty(v) Then
        RoundUpInteger = WorksheetFunction.RoundUp(CDbl(v), 0)
    Else
        RoundUpInteger = 0
    End If
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Function: Nz
' Purpose: Return value or default if Null/Empty
' Parameters:
'   v - Value to check
'   defaultVal - Default value (optional, defaults to 0)
' Returns: Original value or default
'------------------------------------------------------------------------------
Public Function Nz(ByVal v As Variant, Optional ByVal defaultVal As Variant = 0) As Variant
    If IsNull(v) Or IsEmpty(v) Or v = "" Then
        Nz = defaultVal
    Else
        Nz = v
    End If
End Function

'------------------------------------------------------------------------------
' Function: ToDouble
' Purpose: Safely convert variant to Double
' Parameters:
'   v - Value to convert
' Returns: Double value (0 if conversion fails)
'------------------------------------------------------------------------------
Public Function ToDouble(ByVal v As Variant) As Double
    On Error Resume Next
    If IsNumeric(v) And Not IsEmpty(v) And Not IsNull(v) Then
        ToDouble = CDbl(v)
    Else
        ToDouble = 0
    End If
    On Error GoTo 0
End Function
