Attribute VB_Name = "ExcelEconomaExport"
' Delar upp och formaterar om exporterade Excel-filer från Economa

Option Explicit

Public Sub FormateraEconomaBudget()

    Dim objDataWorkbook As Workbook
    Set objDataWorkbook = ActiveWorkbook
    
    Dim strPath As String
    strPath = ActiveWorkbook.Path

    Dim strSearchTerm As String
    strSearchTerm = "ANSVAR"
    
    Dim foundCell As Range
    Set foundCell = Nothing
    
    Dim intCurrentRow As Integer
    Dim intNextRow As Integer
    intCurrentRow = 1
    intNextRow = 0
    
    Dim strRange As String
    
    Dim intNumExports As Integer
    intNumExports = 0
    
    ' Leta upp första förekomst av ANSVAR och sätt första raden. Hantera att det eventuellt saknas
    strRange = "A" & intCurrentRow & ":A10000"
    Set foundCell = ActiveSheet.Range(strRange).Find(What:=strSearchTerm)
    If foundCell Is Nothing Then
        ' Ordet ANSVAR saknas i första kolumnen, fel på bladet
        Call MsgBox("Ordet ANSVAR saknas i första kolumnen. Kontrollera att du är i rätt Excelblad", vbOKOnly, "Budgetformatering Förskola")
        Exit Sub
    Else
        intCurrentRow = foundCell.Row
    End If
    
    ' Sök igenom bladet efter resultat och exportera
    Do While Not foundCell Is Nothing
        
        ' Sök nästa rad med ANSVAR
        strRange = "A" & intCurrentRow + 1 & ":A10000"
        Set foundCell = objDataWorkbook.ActiveSheet.Range(strRange).Find(What:=strSearchTerm)
        If foundCell Is Nothing Then
            ' Ingen nästa cell, sista enheten. Sök sista raden i boken
            intNextRow = Cells(Rows.Count, 1).End(xlUp).Row
        Else
            ' Cell hittad, ange raden för cellen
            intNextRow = foundCell.Row
        End If
        
        Call CopyToNewWorkbook(objDataWorkbook, intCurrentRow + 1, intNextRow - 1, strPath)
        intNumExports = intNumExports + 1
        intCurrentRow = intNextRow
        
        If foundCell Is Nothing Then
            Exit Do
        End If
        
        If intNumExports > 50 Then
            Exit Sub
        End If
    
    Loop
    
    Call MsgBox("Exporten är klar" & vbCrLf & "Antal exporterade enheter: " & intNumExports, vbOKOnly, "Budgetformatering Förskola")
    
End Sub

Public Sub FormateraEconomaTransaktioner()
        
    Dim objTransSheet As Worksheet
    Set objTransSheet = ActiveSheet
    
    Dim strPath As String
    'strPath = "C:\Users\andreas.pettersson\Desktop\Kristina B\Test"
    strPath = ActiveWorkbook.Path
    
    Dim objLookupSheet As Worksheet
    Set objLookupSheet = ActiveWorkbook.Sheets.Add
    
    ' Slå upp de unika ansvaren ur transaktionslistan till ett nytt blad
    ' Funktionen AdvancedFilter förutsätter/antar att sökområdet innehåller en rubrik. Rubeiken ska alltså inkluderas i sökningen.
    Range(objTransSheet.Name & "!D:D").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range(objLookupSheet.Name & "!A1"), Unique:=True
    
    ' Hitta sista raden i det nya bladet
    Dim intLastLookupRow As Integer
    intLastLookupRow = objLookupSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Hitta sista raden i transaktionsbladet
    Dim intLastTransRow As Integer
    intLastTransRow = objTransSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim strAnsvar As String
    Dim strExportName As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim intNumExports As Integer
    intNumExports = 0
    
    Dim objExportWorkbook As Workbook
    Dim objExportWorkSheet As Worksheet
    
    ' Loopa igenom de unika ansvaren
    For i = 2 To intLastLookupRow
        ' Slå upp det aktuella ansvaret
        strAnsvar = objLookupSheet.Cells(i, 1).Value
        strExportName = Left(strAnsvar, 6)
        
        ' Skapa ny arbetsbok
        Set objExportWorkbook = Workbooks.Add
        objExportWorkbook.Sheets(1).Name = strExportName
        Set objExportWorkSheet = objExportWorkbook.Sheets(strExportName)
        objTransSheet.Rows(1).Copy Destination:=objExportWorkSheet.Range("A1")
        
        k = 2
        
        ' Loopa igenom transaktionslistan. Kopiera alla rader som matchar till den nya arbetsboken
        For j = 2 To intLastTransRow
            If objTransSheet.Cells(j, 4) = strAnsvar Then
                objTransSheet.Rows(j).Copy Destination:=objExportWorkSheet.Range("A" & k)
                k = k + 1
            End If
        Next j
        
        Call FormateraTransaktioner(objExportWorkSheet)
        
        objExportWorkbook.SaveAs (strPath & "\" & strExportName & " - Transaktioner")
        objExportWorkbook.Close
        
        intNumExports = intNumExports + 1
        
    Next i
    
    ' Ta bort bladet med uppslag
    Application.DisplayAlerts = False
    objLookupSheet.Delete
    Application.DisplayAlerts = True
    
    Call MsgBox("Exporten är klar" & vbCrLf & "Antal exporterade enheter: " & intNumExports, vbOKOnly, "Transaktionsformatering Förskola")
    
End Sub

Private Sub CopyToNewWorkbook(objWorkbook As Workbook, intFirstRow As Integer, intLastRow As Integer, strTargetPath As String)

    ' Öppnar nytt kalkylblad
    Dim newWorkbook As Workbook
    Set newWorkbook = Workbooks.Add
    
    'Kopierar rubriker
    objWorkbook.Activate
    Range("A1:G1").Select
    Selection.Copy
    
    ' Klistrar in i nya kalkylbladet
    newWorkbook.Activate
    ActiveSheet.Paste
    
    ' Kopierar budgetrader
    objWorkbook.Activate
    Rows(CStr(intFirstRow) & ":" & CStr(intLastRow)).Select
    Application.CutCopyMode = False
    Selection.Copy
    
    ' Klistrar in budgetrader
    newWorkbook.Activate
    Range("A2").Select
    ActiveSheet.Paste
    
    ' Kastar överflödiga blad
    Application.DisplayAlerts = False
    Sheets("Blad3").Select
    Application.CutCopyMode = False
    ActiveWindow.SelectedSheets.Delete
    Sheets("Blad2").Select
    ActiveWindow.SelectedSheets.Delete
    Application.DisplayAlerts = True
    
    ' Sätter namn på bladet
    Dim strNewName As String
    Sheets("Blad1").Select
    strNewName = Cells(2, 1) & " - " & Cells(2, 2)
    Sheets("Blad1").Name = strNewName
    
    Call FormateraBudget(ActiveSheet)
    
    ' Markerar längst upp till vänster
    Range("A1").Select
    
    Call newWorkbook.SaveAs(strTargetPath & "\" & strNewName)
    
    newWorkbook.Close
    
End Sub

Private Sub FormateraBudget(mySheet As Worksheet)

    ' Formaterar det exporterade bladet
    
    Application.PrintCommunication = False
    With mySheet.PageSetup
        .PrintGridlines = True
        .Orientation = xlLandscape
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    Application.PrintCommunication = True

End Sub

Private Sub FormateraTransaktioner(mySheet As Worksheet)

    ' Formaterar det exporterade bladet
    
    mySheet.Columns("A:A").EntireColumn.AutoFit
    mySheet.Columns("B:B").EntireColumn.AutoFit
    mySheet.Columns("C:C").EntireColumn.AutoFit
    mySheet.Columns("E:E").EntireColumn.AutoFit
    mySheet.Columns("F:F").EntireColumn.AutoFit
    mySheet.Columns("D:D").Select
    With Selection
        .HorizontalAlignment = xlLeft
    End With
    
    Application.PrintCommunication = False
    With mySheet.PageSetup
        .Orientation = xlLandscape
        .PrintGridlines = True
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
        .FitToPagesWide = 1
        .FitToPagesTall = 0
    End With
    Application.PrintCommunication = True
    
End Sub
