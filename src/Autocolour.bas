
Attribute VB_Name = "AutocolourShortcut"

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
            cell.Interior.ColorIndex = xlNone ' No fill for blanks
        End If
    Next cell
End Sub
