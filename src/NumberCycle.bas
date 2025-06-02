
Attribute VB_Name = "FormatShortcuts"
Dim formatIndex As Integer

Sub CtrlShift1_NumberCycle()
    Dim formats As Variant
    formats = Array( _
        "#,##0", _
        "#,##0.0", _
        "#,##0.00", _
        "_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)", _
        "0", _
        "General" _
    )

    formatIndex = (formatIndex + 1) Mod (UBound(formats) + 1)

    Dim cell As Range
    For Each cell In Selection
        If Not cell.HasFormula Then
            cell.NumberFormat = formats(formatIndex)
        End If
    Next cell
End Sub
