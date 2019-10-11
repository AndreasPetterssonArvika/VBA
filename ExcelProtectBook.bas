Attribute VB_Name = "ExcelProtectBook"
Option Explicit

Public Sub ProtectBook()

    ' Funktionen skyddar alla blad i en arbetsbok med samma l�senord
    ' Funktionen visar l�senordet i klartext, v�rt att l�sa n�gon g�ng
    
    ' Funktionen fungerar inte om man markerar mer �n ett arbetsblad. Hantering av det inlagt i form av ett meddelande.
    ' Alternativt kan man plocka bort den felhanteringen och avkommentera de tv� f�ljande raderna som markerar �versta cellen i f�rsta arbetsbladet.
    
    'Sheets(1).Select
    'Range("A1").Select
    
    Dim strPW1 As String
    Dim strPW2 As String
    
    strPW1 = InputBox("Ange ett l�senord:", "Skydda arbetsbok")
    strPW2 = InputBox("Ange l�senordet igen:", "Skydda arbetsbok")
    
    If strPW1 <> strPW2 Then
        Call MsgBox("L�senorden matchar inte varandra. F�rs�k igen.", vbOKOnly, "Skydda arbetsbok")
        Exit Sub
    End If
    
    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets
        On Error Resume Next
        Call ws.Protect(strPW1)
        If Err.Number = 1004 Then
            ' Fel, antagligen �r alla blad i arbetsboken markerade. Meddela och avsluta
            Call MsgBox("Fel n�r arbetsboken skulle l�sas." & vbCrLf & "Oftast beror det p� att mer �n ett arbetsblad �r markerade samtidigt." & vbCrLf & "Markera ett enda blad och f�rs�k igen.", vbOKOnly, "Runtime Error 1004")
            Exit Sub
        ElseIf Err.Number <> 0 Then
            ' Annat fel, l�gg upp meddelanderuta med felkod och -beskrivning.
            Call MsgBox("Fel vid k�rning" & vbCrLf & "Felkod: " & Err.Number & vbCrLf & "Beskrivning: " & Err.Description, vbOKOnly, "Runtime Error: " & Err.Number)
        End If
        Err.Clear
        On Error GoTo 0
    Next
    
    Call MsgBox("Arbetsboken �r skyddad", vbOKOnly, "Skydda arbetsbok")
    
End Sub

Public Sub UnprotectBook()

    ' Funktionen tar bort skyddet fr�n en arbetsbok

    Dim strPW As String
    
    strPW = InputBox("Ange l�senordet:", "Ta bort skydd fr�n arbetsbok")
    
    Dim ws As Worksheet
    Dim boolSuccess As Boolean
    boolSuccess = True
    
    For Each ws In ActiveWorkbook.Worksheets
        On Error Resume Next
        Call ws.Unprotect(strPW)
        
        ' Hantera felaktigt l�senord
        If Err.Number = 1004 Then
            Call MsgBox("Felaktigt l�senord f�r blad " & ws.Name, vbOKOnly, "Ta bort skydd fr�n arbetsbok")
            boolSuccess = False
        ElseIf Err.Number <> 0 Then
            Call MsgBox("Ov�ntat fel." & vbCrLf & "Felkod: " & Err.Number & vbCrLf & "Meddelande: " & Err.Description, vbOKOnly, "Ta bort skydd fr�n arbetsbok")
            Exit Sub
        End If
        
        Err.Clear
        On Error GoTo 0
        
    Next
    
    If boolSuccess Then
        Call MsgBox("Arbetsboken �r uppl�st", vbOKOnly, "Ta bort skydd fr�n arbetsbok")
    Else
        Call MsgBox("Ett eller flera blad kunde inte l�sas upp p� grund av ett felaktigt l�senord.", vbOKOnly, "Ta bort skydd fr�n arbetsbok")
    End If

End Sub


