Attribute VB_Name = "ExcelImporteraCosmicLink"
' Formaterar om och kopierar exporterade CSV-filer fr�n CosmicLink.
' F�ruts�tter att modulen OpenSaveDialog finns med

Option Explicit

Sub ImporteraCosmicLink()

'
' ImporteraCosmicLink Makro
'

    Dim strImportFile As String
    Dim strOriginPath As String     ' Anger mappen d�r filera f�rv�ntas finnas
    Dim strConnectionString As String
    
    strOriginPath = "<PATH>"

    strImportFile = GetNameCsv(strOriginPath, , "H�mta fil")
    strConnectionString = "TEXT;" & strImportFile

   
    ' M�lmapp f�r importen
    Dim strTargetFolder As String
    strTargetFolder = "<PATH>"

    ' Importera det angivna bladet
    With ActiveSheet.QueryTables.Add(Connection:=strConnectionString, Destination:=Range("$A$1"))
    
        .Name = "Link_Test"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 1252
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileCommaDelimiter = True
        .TextFileColumnDataTypes = Array(2, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False

    End With
    
    Dim strFileName As String
    strFileName = Right(strImportFile, Len(strImportFile) - InStrRev(strImportFile, "\"))
    strFileName = Left(strFileName, InStr(strFileName, ".") - 1)
    
    Dim strExportPath As String
    strExportPath = strTargetFolder & "\" & strFileName & ".xlsx"

    ' Sparar filen
    ActiveWorkbook.SaveAs FileName:=strExportPath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

    ' St�nger det bearbetade bladet och �ppnar ett nytt Excel-blad
    ActiveWorkbook.Close
    Call Workbooks.Add

End Sub

