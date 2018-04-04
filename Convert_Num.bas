Attribute VB_Name = "Module1"
Sub Convert_Num()
    For Each xCell In Selection
        Selection.NumberFormat = "0" 'Note: The "0.00" determines the number of decimal places.
        xCell.value = xCell.value
    Next xCell
End Sub
