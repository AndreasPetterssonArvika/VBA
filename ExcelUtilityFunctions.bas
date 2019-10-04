Option Explicit

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

Sub ProtectBook()

    ' Funktionen skyddar alla blad i en arbetsbok med samma lösenord
    ' Funktionen visar lösenordet i klartext, värt att lösa någon gång
    
    ' Funktionen fungerar inte om man markerar mer än ett arbetsblad. Hantering av det inlagt i form av ett meddelande.
    ' Alternativt kan man plocka bort den felhanteringen och avkommentera de två följande raderna som markerar översta cellen i första arbetsbladet.
    
    'Sheets(1).Select
    'Range("A1").Select
    
    Dim strPW1 As String
    Dim strPW2 As String
    
    strPW1 = InputBox("Ange ett lösenord:", "Skydda arbetsbok")
    strPW2 = InputBox("Ange lösenordet igen:", "Skydda arbetsbok")
    
    If strPW1 <> strPW2 Then
        Call MsgBox("Lösenorden matchar inte varandra. Försök igen.", vbOKOnly, "Skydda arbetsbok")
        Exit Sub
    End If
    
    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets
        On Error Resume Next
        Call ws.Protect(strPW1)
        If Err.Number = 1004 Then
            ' Fel, antagligen är alla blad i arbetsboken markerade. Meddela och avsluta
            Call MsgBox("Fel när arbetsboken skulle låsas." & vbCrLf & "Oftast beror det på att mer än ett arbetsblad är markerade samtidigt." & vbCrLf & "Markera ett enda blad och försök igen.", vbOKOnly, "Runtime Error 1004")
            Exit Sub
        ElseIf Err.Number <> 0 Then
            ' Annat fel, lägg upp meddelanderuta med felkod och -beskrivning.
            Call MsgBox("Fel vid körning" & vbCrLf & "Felkod: " & Err.Number & vbCrLf & "Beskrivning: " & Err.Description, vbOKOnly, "Runtime Error: " & Err.Number)
        End If
        Err.Clear
        On Error GoTo 0
    Next
    
    Call MsgBox("Arbetsboken är skyddad", vbOKOnly, "Skydda arbetsbok")
    
End Sub

Sub UnprotectBook()

    ' Funktionen tar bort skyddet från en arbetsbok

    Dim strPW As String
    
    strPW = InputBox("Ange lösenordet:", "Ta bort skydd från arbetsbok")
    
    Dim ws As Worksheet
    Dim boolSuccess As Boolean
    boolSuccess = True
    
    For Each ws In ActiveWorkbook.Worksheets
        On Error Resume Next
        Call ws.Unprotect(strPW)
        
        ' Hantera felaktigt lösenord
        If Err.Number = 1004 Then
            Call MsgBox("Felaktigt lösenord för blad " & ws.Name, vbOKOnly, "Ta bort skydd från arbetsbok")
            boolSuccess = False
        ElseIf Err.Number <> 0 Then
            Call MsgBox("Oväntat fel." & vbCrLf & "Felkod: " & Err.Number & vbCrLf & "Meddelande: " & Err.Description, vbOKOnly, "Ta bort skydd från arbetsbok")
            Exit Sub
        End If
        
        Err.Clear
        On Error GoTo 0
        
    Next
    
    If boolSuccess Then
        Call MsgBox("Arbetsboken är upplåst", vbOKOnly, "Ta bort skydd från arbetsbok")
    Else
        Call MsgBox("Ett eller flera blad kunde inte låsas upp på grund av ett felaktigt lösenord.", vbOKOnly, "Ta bort skydd från arbetsbok")
    End If

End Sub

