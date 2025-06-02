
Attribute VB_Name = "PercentFormatShortcuts"
Dim percentFormatIndex As Integer

Sub CtrlShift5_PercentCycle()
    Dim formats As Variant
    formats = Array("0%", "0.0%", "0.00%")

    percentFormatIndex = (percentFormatIndex + 1) Mod (UBound(formats) + 1)

    Dim cell As Range
    For Each cell In Selection
        If IsNumeric(cell.Value) Then
            ' Apply only if current format contains "%" or is numeric
            If InStr(cell.NumberFormat, "%") > 0 Or cell.NumberFormat = "General" Then
                cell.NumberFormat = formats(percentFormatIndex)
            End If
        End If
    Next cell
End Sub
