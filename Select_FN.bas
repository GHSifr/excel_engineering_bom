Attribute VB_Name = "Module3"
Sub Select_FN()
Dim LR As Long, cell As Range, rng As Range
With Sheets("Sheet1")
    LR = .Range("FN" & Rows.Count).End(xlUp).Row
    For Each cell In .Range("A14:A1000" & LR)
        If cell.value <> "" Then
            If rng Is Nothing Then
                Set rng = cell
            Else
                Set rng = Union(rng, cell)
            End If
        End If
    Next cell
    rng.Select
End With
End Sub
