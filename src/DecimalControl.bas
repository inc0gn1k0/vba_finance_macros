
Attribute VB_Name = "DecimalControlShortcut"

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
    Dim base As String
    Dim decimals As Integer
    Dim newFormat As String

    ' Extract current decimal places
    If InStr(fmt, ".") > 0 Then
        base = Split(fmt, ".")(0)
        decimals = Len(Split(fmt, ".")(1))
    Else
        base = fmt
        decimals = 0
    End If

    ' Adjust decimal places
    decimals = Application.Max(0, decimals + delta)

    newFormat = base & IIf(decimals > 0, "." & String(decimals, "0"), "")
    AdjustDecimals = newFormat
End Function
