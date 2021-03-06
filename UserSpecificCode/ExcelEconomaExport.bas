Attribute VB_Name = "ExcelEconomaExport"
' Delar upp och formaterar om exporterade Excel-filer fr�n Economa.
' F�ruts�tter att modulen ExcelUtilityFunctions finns.

Option Explicit

' Visas bl a n�r man f�rs�ker index f�r Excelblad som inte finns
Public Const INDEX_OUT_OF_RANGE = 9

Public Sub FormateraEconomaBudget()

    ' Set start time for timer, only needed for performance testing
    'Dim dTime As Double
    'dTime = MicroTimer
    
    TurnOffStuff

    Dim shSource As Worksheet
    Set shSource = ActiveWorkbook.Sheets(1)
    
    Dim objNewWorkbook As Workbook
    Dim shOutput As Worksheet
    
    Dim strPath As String
    strPath = VFD_GetFolderPath(ActiveWorkbook.Path, "V�lj m�lmapp", "V�lj")

    Dim strSearchTerm As String
    strSearchTerm = "ANSVAR"
    
    Dim foundCell As Range
    Set foundCell = Nothing
    
    Dim intCurrentRow As Integer
    Dim intNextRow As Integer
    Dim intLastCopyRow As Integer
    Dim intNumRowsToCopy As Integer
    intCurrentRow = 1
    intNextRow = 0
    intLastCopyRow = 0
    intNumRowsToCopy = 0
    
    Dim intNumExports As Integer
    intNumExports = 0
    
    ' Leta upp f�rsta f�rekomst av ANSVAR och s�tt f�rsta raden. Hantera att det eventuellt saknas
    Set foundCell = shSource.Range(shSource.Cells(1, 1), shSource.Cells(10000, 1)).Find(What:=strSearchTerm)
    If foundCell Is Nothing Then
        ' Ordet ANSVAR saknas helt i f�rsta kolumnen, fel p� bladet
        Call MsgBox("Ordet ANSVAR saknas i f�rsta kolumnen. Kontrollera att du �r i r�tt Excelblad", vbOKOnly, "Budgetformatering F�rskola")
        Exit Sub
    Else
        intCurrentRow = foundCell.Row
    End If
    
    ' S�k igenom bladet efter resultat och exportera
    Do While Not foundCell Is Nothing
        
        ' S�k n�sta rad med ANSVAR.
        ' S�ker de f�rsta 10000 raderna vilket borde r�cka med marginal eftersom k�lldatabladet idag har ca 650 rader
        Set foundCell = shSource.Range(shSource.Cells(intCurrentRow + 1, 1), shSource.Cells(10000, 1)).Find(What:=strSearchTerm)
        If foundCell Is Nothing Then
            ' Ingen n�sta cell, sista enheten. S�k sista raden i boken och dra av en rad eftersom sista raden inneh�ller totalerna f�r hela bladet
            intLastCopyRow = Cells(Rows.Count, 1).End(xlUp).Row - 1
        Else
            ' Cell med ANSVAR hittad, ange raden f�r cellen och l�gg �ven in sista raden som ska kopieras vilket �r tv� rader ovanf�r
            intNextRow = foundCell.Row
            intLastCopyRow = intNextRow - 2
        End If
        
        ' R�kna ut antalet rader som ska kopieras
        intNumRowsToCopy = intLastCopyRow - intCurrentRow + 1
        
        Set objNewWorkbook = Workbooks.Add
        Application.DisplayAlerts = False
        
        ' Fel n�r blad 2 och 3 ska tas bort ur den nya Excelboken
		' Detta beror p� att Office 365 bara har ett blad i nya arbetsb�cker.
		On Error Resume Next
        objNewWorkbook.Sheets(3).Delete
        If Err.Number = INDEX_OUT_OF_RANGE Then
            ' Fliken saknas, normalt i Office 365. Rensa bara felkoden.
            Err.Clear
        ElseIf Err.Number <> 0 Then
            ' N�got annat fel har intr�ffat, meddela i dialogruta
            Call MsgBox("Fel vid k�rning" & vbCrLf & "Felkod: " & Err.Number & vbCrLf & "Felmeddelanden: " & Err.Description, vbOKOnly, "Fel vid k�rning")
            Exit Sub
        End If
        On Error GoTo 0

        On Error Resume Next
        objNewWorkbook.Sheets(2).Delete
        If Err.Number = INDEX_OUT_OF_RANGE Then
            ' Fliken saknas, normalt i Office 365. Rensa bara felkoden.
            Err.Clear
        ElseIf Err.Number <> 0 Then
            ' N�got annat fel har intr�ffat, meddela i dialogruta
            Call MsgBox("Fel vid k�rning" & vbCrLf & "Felkod: " & Err.Number & vbCrLf & "Felmeddelanden: " & Err.Description, vbOKOnly, "Fel vid k�rning")
            Exit Sub
        End If
        On Error GoTo 0
        
        
        Application.DisplayAlerts = True
        Set shOutput = objNewWorkbook.Sheets(1)
        
        ' Formatera utdatabladets f�rsta kolumn som text
        shOutput.Columns("A:A").EntireColumn.NumberFormat = "@"
        
        ' Kopiera bladets rubriker till utdatabladet
        shOutput.Range("A1:G1").Value = shSource.Range("A1:G1").Value
        
        ' Kopiera den aktuella budgeten till utdatabladet
        shOutput.Range(shOutput.Cells(2, 1), shOutput.Cells(intNumRowsToCopy + 1, 7)).Value = shSource.Range(shSource.Cells(intCurrentRow, 1), shSource.Cells(intLastCopyRow, 7)).Value
        
        intNumExports = intNumExports + 1
        intCurrentRow = intNextRow
        
        ' S�tter namn p� bladet
        shOutput.Name = shOutput.Cells(3, 1) & " - " & shOutput.Cells(3, 2)
        
        ' Formatera utdatabladet
        Call FormateraBudget(shOutput)
        
        ' Markerar l�ngst upp till v�nster
        Range("A1").Select
        
        Call objNewWorkbook.SaveAs(strPath & "\" & shOutput.Name)
        
        
        objNewWorkbook.Close
        Set objNewWorkbook = Nothing
        
        
        If foundCell Is Nothing Then
            Exit Do
        End If
        
        If intNumExports > 100 Then
            Exit Sub
        End If
    
    Loop
    
    TurnOnStuff
    
    ' Print the result of the timer
    'Debug.Print vbCrLf & "Time is: " & (MicroTimer - dTime) * 1000
    
    Call MsgBox("Exporten �r klar" & vbCrLf & "Antal exporterade enheter: " & intNumExports, vbOKOnly, "Budgetformatering F�rskola")
    
End Sub

Public Sub FormateraEconomaTransaktioner()

    TurnOffStuff
        
    Dim objTransSheet As Worksheet
    Set objTransSheet = ActiveSheet
    
    Dim strPath As String
    strPath = VFD_GetFolderPath(ActiveWorkbook.Path, "V�lj m�lmapp", "V�lj")
    
    Dim objLookupSheet As Worksheet
    Set objLookupSheet = ActiveWorkbook.Sheets.Add
    
    ' Sl� upp de unika ansvaren ur transaktionslistan till ett nytt blad
    ' Funktionen AdvancedFilter f�ruts�tter/antar att s�komr�det inneh�ller en rubrik. Rubriken ska allts� inkluderas i s�kningen.
    Range(objTransSheet.Name & "!E:E").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range(objLookupSheet.Name & "!A1"), Unique:=True
    
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
        ' Sl� upp det aktuella ansvaret
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
            If objTransSheet.Cells(j, 5) = strAnsvar Then
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
    
    TurnOnStuff
    
    Call MsgBox("Exporten �r klar" & vbCrLf & "Antal exporterade enheter: " & intNumExports, vbOKOnly, "Transaktionsformatering F�rskola")
    
End Sub

Private Sub FormateraBudget(mySheet As Worksheet)

    ' Formaterar det exporterade bladet
    
    ' H�gerst�ll kolumnerna C till G och l�gg text h�gst upp i cellen
    Columns("C:G").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlTop
    End With
    
    ' V�nsterst�ll kolumnerna A och B
    Columns("A:B").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
    End With
    
    ' S�tt rubriken i B1 uptill i cellen
    Range("B1").Select
    With Selection
        .VerticalAlignment = xlTop
    End With

    ' Justera kolumnbredder
    Columns("A:A").ColumnWidth = 20
    Columns("B:B").ColumnWidth = 20
    Columns("C:C").ColumnWidth = 20
    Columns("D:D").ColumnWidth = 20
    Columns("E:E").ColumnWidth = 20
    Columns("F:F").ColumnWidth = 20
    Columns("G:G").ColumnWidth = 20
    
    ' S�tter utskriftsformat
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
    mySheet.Columns("D:D").HorizontalAlignment = xlLeft
    
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

