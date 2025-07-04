
Attribute VB_Name = "FinancialModelingShortcuts"
' ===============================
' Module-level Variables
' ===============================
Dim formatIndex As Integer
Dim dateFormatIndex As Integer
Dim percentFormatIndex As Integer
Dim multipleFormatIndex As Integer
Dim borderCycleIndex As Integer

' ===============================
' Ctrl+Shift+1: Number Format Cycle
' ===============================
Sub CtrlShift1_NumberCycle()
    Dim formats As Variant
    formats = Array("#,##0", "#,##0.0", "#,##0.00", "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)", "0", "General")
    formatIndex = (formatIndex + 1) Mod (UBound(formats) + 1)
    Dim cell As Range
    For Each cell In Selection
        If Not cell.HasFormula Then
            cell.NumberFormat = formats(formatIndex)
        End If
    Next cell
End Sub

' ===============================
' Ctrl+Shift+2: Date Format Cycle
' ===============================
Sub CtrlShift2_DateCycle()
    Dim formats As Variant
    formats = Array("dd/mm/yyyy", "dd-mmm-yyyy", "mmm-yy", "mmmm dd, yyyy", "mm/dd/yyyy", "yyyy-mm-dd")
    dateFormatIndex = (dateFormatIndex + 1) Mod (UBound(formats) + 1)
    Dim cell As Range
    For Each cell In Selection
        If IsDate(cell.Value) Then
            cell.NumberFormat = formats(dateFormatIndex)
        End If
    Next cell
End Sub

' ===============================
' Ctrl+Shift+5: Percent Format Cycle
' ===============================
Sub CtrlShift5_PercentCycle()
    Dim formats As Variant
    formats = Array("0%", "0.0%", "0.00%")
    percentFormatIndex = (percentFormatIndex + 1) Mod (UBound(formats) + 1)
    Dim cell As Range
    For Each cell In Selection
        If IsNumeric(cell.Value) Then
            If InStr(cell.NumberFormat, "%") > 0 Or cell.NumberFormat = "General" Then
                cell.NumberFormat = formats(percentFormatIndex)
            End If
        End If
    Next cell
End Sub

' ===============================
' Ctrl+Shift+8: Financial Multiples Cycle
' ===============================
Sub CtrlShift8_MultipleCycle()
    Dim formats As Variant
    formats = Array("#,##0", "0.0""x""", "0.0,""K""", "0.0,,""M""", "0.00,,,""B""")
    multipleFormatIndex = (multipleFormatIndex + 1) Mod (UBound(formats) + 1)
    Dim cell As Range
    For Each cell In Selection
        If IsNumeric(cell.Value) Then
            cell.NumberFormat = formats(multipleFormatIndex)
        End If
    Next cell
End Sub

' ===============================
' Ctrl+Alt+A: Autocolour
' ===============================
Sub CtrlAltA_Autocolour()
    Dim cell As Range
    For Each cell In Selection
        If cell.HasFormula Then
            cell.Interior.Color = RGB(198, 239, 206) ' Light green
        ElseIf IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            cell.Interior.Color = RGB(222, 235, 247) ' Light blue
        ElseIf VarType(cell.Value) = vbString And Trim(cell.Value) <> "" Then
            cell.Interior.Color = RGB(217, 217, 217) ' Light gray
        Else
            cell.Interior.ColorIndex = xlNone
        End If
    Next cell
End Sub

' ===============================
' Ctrl+Alt+Shift+Arrow: Border Cycle
' ===============================
Sub CtrlAltShift_BorderCycle()
    Dim cell As Range
    Dim borders As Variant
    borders = Array("None", "Bottom", "Top", "All")
    borderCycleIndex = (borderCycleIndex + 1) Mod (UBound(borders) + 1)
    For Each cell In Selection
        With cell.Borders
            .LineStyle = xlNone
        End With
        Select Case borders(borderCycleIndex)
            Case "Bottom"
                With cell.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
            Case "Top"
                With cell.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
            Case "All"
                With cell.Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
        End Select
    Next cell
End Sub

' ===============================
' Ctrl+Shift+N: Switch to Negative
' ===============================
Sub CtrlShiftN_SwitchToNegative()
    Dim cell As Range
    For Each cell In Selection
        If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            cell.Value = -1 * cell.Value
        End If
    Next cell
End Sub

' ===============================
' Ctrl+< / Ctrl+>: Decimal Adjustment
' ===============================
Sub CtrlComma_DecreaseDecimal()
    Dim cell As Range
    For Each cell In Selection
        If IsNumeric(cell.Value) Then
            cell.NumberFormat = AdjustDecimals(cell.NumberFormat, -1)
        End If
    Next cell
End Sub

Sub CtrlPeriod_IncreaseDecimal()
    Dim cell As Range
    For Each cell In Selection
        If IsNumeric(cell.Value) Then
            cell.NumberFormat = AdjustDecimals(cell.NumberFormat, 1)
        End If
    Next cell
End Sub

Function AdjustDecimals(fmt As String, delta As Integer) As String
    Dim cleanedFmt As String
    Dim base As String
    Dim decimals As Integer
    Dim i As Integer
    
    ' Strip formatting to a core decimal pattern
    If fmt = "General" Or fmt = "" Then
        base = "0"
        decimals = 0
    ElseIf InStr(fmt, ".") > 0 Then
        base = Split(fmt, ".")(0)
        decimals = Len(Split(fmt, ".")(1))
    Else
        base = fmt
        decimals = 0
    End If

    decimals = Application.Max(0, decimals + delta)
    AdjustDecimals = base & IIf(decimals > 0, "." & String(decimals, "0"), "")
End Function