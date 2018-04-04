Private Sub Workbook_BeforePrint(Cancel As Boolean)

' DH 2013/02/21
' DH - I may implement this idea at a later date.
'
'    With ActiveSheet.PageSetup.CenterHeader = _
'        Format(Worksheets("PL").Range("C8").Value)
'    End With
'
'    With ActiveSheet.PageSetup.LeftFooter = _
'    "Date Printed:   " & Format(Date, "dd-mmm-yyyy") & vbLf & _
'    "Time Printed:   " & Format(Time, "hhmm") & " hrs"
'    End With

    
    Call UpdateHeader
    Call Select_All
    With Selection.Font
        .Name = "Arial"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        '.MergeCells = False
    End With
    
    Range("A14:B1000").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    Call Select_FN
    With Selection
    Call Convert_Num
    
    End With
    

    'Range("A14:F1000").Select
    'Selection.Rows.AutoFit
    Range("A14").Select

End Sub


Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, _
        Cancel As Boolean)
    
    Call UpdateHeader
    Call Select_All
    With Selection.Font
        .Name = "Arial"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        '.MergeCells = False
    End With
    
    Range("A14:B1000").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    Call Select_FN
    With Selection
    Call Convert_Num
    
    End With
    

    'Range("A14:F1000").Select
    'Selection.Rows.AutoFit

    Range("A14").Select

End Sub