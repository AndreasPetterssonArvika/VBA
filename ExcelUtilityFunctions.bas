Attribute VB_Name = "ExcelUtilityFunctions"
Option Explicit
Option Private Module

Private Const ACTION_BUTTON_PRESSED = -1
Private Const CANCEL_BUTTON_PRESSED = 0

Private Declare Function getFrequency Lib "kernel32" _
Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
Private Declare Function getTickCount Lib "kernel32" _
Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long

Public Sub TurnOffStuff()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
End Sub

Public Sub TurnOnStuff()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Public Function MicroTimer() As Double

' Returns seconds.
Dim cyTicks1 As Currency
Static cyFrequency As Currency
' Initialize MicroTimer
MicroTimer = 0
' Get frequency.
If cyFrequency = 0 Then getFrequency cyFrequency
' Get ticks.
getTickCount cyTicks1
' Seconds = Ticks (or counts) divided by Frequency
If cyFrequency Then MicroTimer = cyTicks1 / cyFrequency

End Function

' Följande funktioner innehåller kod för att hämta sökvägar till mapper eller när filer ska öppnas eller sparas.
' Den innehåller ett antal standardfilter för att t ex öppna vissa filtyper

' https://docs.microsoft.com/en-us/office/vba/api/office.filedialog

' FileDialog finns i fyra varianter.
' - msoFileDialogFilePicker, hämtar en sökväg till en fil. Kan förses med filter
' - msoFileDialogFolderPicker, hämtar en sökväg till en mapp.
' - msoFileDialogOpen, öppnar en fil som applikationen kan hantera.
' - msoFileDialogSaveAs, sparar en fil i ett format som applikationen kan hantera.

Public Function VFD_GetFolderPath(Optional strInitialFileName As String = vbNullString, Optional strTitle As String = "Bläddra", Optional strButtonName As String = "OK") As String

    ' Låter användaren bläddra fram sökvägen till en mapp och returnerar sökvägen eller vbNullString
    ' Vid upprepade anrop kommer dialogen ihåg mappen om den inte anges explicit
    
    Dim f As FileDialog
    Dim lngReturn As Long
    
    Set f = Application.FileDialog(msoFileDialogFolderPicker)
    
    If strInitialFileName <> vbNullString Then
        f.InitialFileName = strInitialFileName
    End If
    
    f.Title = strTitle
    f.ButtonName = strButtonName
    
    lngReturn = f.Show
    
    If lngReturn = ACTION_BUTTON_PRESSED Then
        VFD_GetFolderPath = f.SelectedItems(1)
        Exit Function
    End If
    
    VFD_GetFolderPath = vbNullString
    
End Function

Public Function VFD_GetTextFileName(Optional strInitialFileName As String = vbNullString, Optional strTitle As String = "Bläddra", Optional strButtonName As String = "Öppna") As String

    ' Öppnar en dialogruta för att hämta sökvägen till en textfil
    
    Dim f As FileDialog
    Dim lngReturn As Long
    Set f = Application.FileDialog(msoFileDialogFilePicker)
    
    If strInitialFileName <> vbNullString Then
        f.InitialFileName = strInitialFileName
    End If
    
    f.Title = strTitle
    f.ButtonName = strButtonName
    
    Call f.Filters.Clear
    Call f.Filters.Add("Textfiler", "*.txt", 1)
    Call f.Filters.Add("Alla filer", "*.*", 2)
    f.FilterIndex = 1
    
    lngReturn = f.Show
    
    If lngReturn = ACTION_BUTTON_PRESSED Then
        VFD_GetTextFileName = f.SelectedItems(1)
        Exit Function
    End If
    
    VFD_GetTextFileName = vbNullString
    
End Function

Public Function VFD_GetTextFileNameSem(Optional strInitialFileName As String = vbNullString, Optional strTitle As String = "Bläddra", Optional strButtonName As String = "Öppna") As String

    ' Öppnar en dialogruta för att hämta sökvägen till en textfil
    
    Dim f As FileDialog
    Dim lngReturn As Long
    Set f = Application.FileDialog(msoFileDialogFilePicker)
    
    If strInitialFileName <> vbNullString Then
        f.InitialFileName = strInitialFileName
    End If
    
    f.Title = strTitle
    f.ButtonName = strButtonName
    
    Call f.Filters.Clear
    Call f.Filters.Add("Semikolonseparerade filer", "*.sem", 1)
    Call f.Filters.Add("Alla filer", "*.*", 2)
    f.FilterIndex = 1
    
    lngReturn = f.Show
    
    If lngReturn = ACTION_BUTTON_PRESSED Then
        VFD_GetTextFileNameSem = f.SelectedItems(1)
        Exit Function
    End If
    
    VFD_GetTextFileNameSem = vbNullString
    
End Function

Public Function VFD_GetExcelFileName(Optional strInitialFileName As String = vbNullString, Optional strTitle As String = "Bläddra", Optional strButtonName As String = "Öppna") As String

    ' Öppnar en dialogruta för att spara en Excel-fil
    Dim f As FileDialog
    Dim lngReturn As Long
    Set f = Application.FileDialog(msoFileDialogFilePicker)
    
    If strInitialFileName <> vbNullString Then
        f.InitialFileName = strInitialFileName
    End If
    
    f.Title = strTitle
    f.ButtonName = strButtonName
    
    Call f.Filters.Clear
    Call f.Filters.Add("Excel-filer", "*.xlsx,*.xls", 1)
    Call f.Filters.Add("Alla filer", "*.*", 2)
    f.FilterIndex = 1
    
    lngReturn = f.Show
    
    If lngReturn = ACTION_BUTTON_PRESSED Then
        VFD_GetExcelFileName = f.SelectedItems(1)
        Exit Function
    End If
    
    VFD_GetExcelFileName = vbNullString
    
End Function

