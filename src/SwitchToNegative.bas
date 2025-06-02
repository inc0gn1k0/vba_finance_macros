
Attribute VB_Name = "NegativeSwitcherShortcut"

Sub CtrlShiftN_SwitchToNegative()
    Dim cell As Range
    For Each cell In Selection
        If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            cell.Value = -1 * cell.Value
        End If
    Next cell
End Sub
