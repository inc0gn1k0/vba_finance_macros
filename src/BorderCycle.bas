
Attribute VB_Name = "BorderCycleShortcut"
Dim borderCycleIndex As Integer

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
                cell.Borders(xlEdgeBottom).LineStyle = xlContinuous
                cell.Borders(xlEdgeBottom).Weight = xlThin
            Case "Top"
                cell.Borders(xlEdgeTop).LineStyle = xlContinuous
                cell.Borders(xlEdgeTop).Weight = xlThin
            Case "All"
                cell.Borders(xlEdgeTop).LineStyle = xlContinuous
                cell.Borders(xlEdgeBottom).LineStyle = xlContinuous
                cell.Borders(xlEdgeLeft).LineStyle = xlContinuous
                cell.Borders(xlEdgeRight).LineStyle = xlContinuous
                cell.Borders(xlEdgeTop).Weight = xlThin
                cell.Borders(xlEdgeBottom).Weight = xlThin
                cell.Borders(xlEdgeLeft).Weight = xlThin
                cell.Borders(xlEdgeRight).Weight = xlThin
        End Select
    Next cell
End Sub
