
Attribute VB_Name = "DateFormatShortcuts"
Dim dateFormatIndex As Integer

Sub CtrlShift2_DateCycle()
    Dim formats As Variant
    formats = Array( _
        "dd/mm/yyyy", _
        "dd-mmm-yyyy", _
        "mmm-yy", _
        "mmmm dd, yyyy", _
        "mm/dd/yyyy", _
        "yyyy-mm-dd" _
    )

    dateFormatIndex = (dateFormatIndex + 1) Mod (UBound(formats) + 1)

    Dim cell As Range
    For Each cell In Selection
        If IsDate(cell.Value) Then
            cell.NumberFormat = formats(dateFormatIndex)
        End If
    Next cell
End Sub
