Attribute VB_Name = "ExcelEconomaExport"
' Formaterar exporterade Excel-blad från Economa

Private Sub EconomaBudget()
'
' EconomaBudget Makro
' Formaterar Excelexport  av budget ur Economa
'
    
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintGridlines = True
        .Orientation = xlLandscape
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    Application.PrintCommunication = True
    
    SaveFile = GetNameExcel(, , "Spara")
    Debug.Print SaveFile

End Sub

Private Sub EconomaTransaktioner()
'
' EconomaTransaktioner Makro
' Formaterar exporterade transaktioner från Economa
'

'
    Columns("A:A").EntireColumn.AutoFit
    Columns("B:B").EntireColumn.AutoFit
    Columns("C:C").EntireColumn.AutoFit
    Columns("E:E").EntireColumn.AutoFit
    Columns("F:F").EntireColumn.AutoFit
    Columns("D:D").Select
    With Selection
        .HorizontalAlignment = xlLeft
    End With
    
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .Orientation = xlLandscape
        .PrintGridlines = True
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
        .FitToPagesWide = 1
        .FitToPagesTall = 0
    End With
    Application.PrintCommunication = True
    
End Sub

