
Attribute VB_Name = "FinancialMultipleShortcuts"
Dim multipleFormatIndex As Integer

Sub CtrlShift8_MultipleCycle()
    Dim formats As Variant
    formats = Array( _
        "#,##0", _
        "0.0"x"", _
        "0.0,"K"", _
        "0.0,," M"", _
        "0.00,,," B"" _
    )

    multipleFormatIndex = (multipleFormatIndex + 1) Mod (UBound(formats) + 1)

    Dim cell As Range
    For Each cell In Selection
        If IsNumeric(cell.Value) Then
            cell.NumberFormat = formats(multipleFormatIndex)
        End If
    Next cell
End Sub
