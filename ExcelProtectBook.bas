Attribute VB_Name = "ExcelProtectBook"
Option Explicit

Public Sub ProtectBook()

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

Public Sub UnprotectBook()

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


