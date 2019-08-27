Public Function GAM_Singular(strMail As String, <parametrar>, Optional boolElevatedShell As Boolean = True) As Integer

    ' Skelettfunktion för att köra ett GAM-kommando för en användare, grupp eller likande.
	
	' boolElevatedShell anger om kommandot ska köras med förhöjda priviiegier.
	' Detta krävs  för att kunna uppdatera filer relaterade till Googles
	' säkerhetslösning.
	' Vid körning av enstaka användare ska det vara True
	' Vid körning av listor räcker det med att första körningen i listan är True
    
    Dim intReturn As Integer
    intReturn = CLEAN_EXIT
    
    Dim strFunctionName As String
    Dim strMessage As String
    
    strFunctionName = "GAM_Singular"
    strMessage = vbNullString
    
	' Skapa parameter-sträng till GAM-kommandot
    Dim strParams As String
    strParams = <skapa parametersträng>
    
	' Kör kommandot.
    If boolElevatedShell Then
        On Error Resume Next
        Call System_RunInElevatedShell("gam.exe", strParams)
        If Err.Number <> CLEAN_EXIT Then
            strMessage = "Fel vid körning av GAM-kommando. Kontrollera att GAM är installerat på datorn." & vbCrLf & "Felkod: " & Err.Number & "Felmeddelande: " & Err.Description
            Call MsgBox(strMessage, vbOKOnly, strFunctionName)
            Call AddLogEntry(LOG_TYPE_ERROR, strFunctionName, strMessage)
            intReturn = BAD_EXIT
        End If
        Err.Clear
        On Error GoTo 0
    Else
        On Error Resume Next
        Call Shell("gam " & strParams)
        If Err.Number <> CLEAN_EXIT Then
            strMessage = "Fel vid körning av GAM-kommando. Kontrollera att GAM är installerat på datorn." & vbCrLf & "Felkod: " & Err.Number & "Felmeddelande: " & Err.Description
            Call MsgBox(strMessage, vbOKOnly, strFunctionName)
            Call AddLogEntry(LOG_TYPE_ERROR, strFunctionName, strMessage)
            intReturn = BAD_EXIT
        End If
        Err.Clear
        On Error GoTo 0
    End If
    
    GAM_Singular = intReturn

End Function

Public Function GAM_Plural(strDataSource As String, <parametrar>) As Integer

    ' Skelettfunktion för att köra ett GAM-kommando för en lista av användare, grupper eller liknande
    
    Dim intReturn As Integer
    intReturn = CLEAN_EXIT
    
    Dim strFunctionName As String
    Dim strMessage As String
    
    strFunctionName = "GAM_Plural"
    strMessage = vbNullString
    
    ' Kontrollera att alla nödvändiga fält finns i datakällan
    Dim strReqFields As String
    strReqFields = "mail"
    
    intReturn = DB_CheckRequiredFieldsDataSource(strDataSource, strReqFields)
    If intReturn <> CLEAN_EXIT Then
        ' Datakällan innehåller inte nödvändiga fält, logga och avsluta
        strMessage = "Datakällan innehåller inte det nödvändiga fältet " & strReqFields & "."
        Call AddLogEntry(LOG_TYPE_ERROR, strFunctionName, strMessage)
        intReturn = BAD_EXIT
        GAM_Plural = intReturn
        Exit Function
    End If
    
    ' Slå upp alla <mål> till recordset
    Dim strSQL As String
    Dim rst<Namn> As Recordset
    
    strSQL = "SELECT " & strReqFields & " FROM " & strDataSource
    
    Set rst<Namn> = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset)
    
    ' Variabel för att ange om behörigheten ska sättas med förhöjd behörighet vid första körningen av GAM_Singular
    Dim boolElevated As Boolean
    boolElevated = True
    
    With rst<Namn>
        While Not .EOF
            intReturn = GAM_Singular(!mail, <parametrar>, boolElevated)
            If intReturn <> CLEAN_EXIT Then
                ' Fel vid körning av kommando, logga och avsluta
                strMessage = "<felmeddelande> " & !mail & "."
                Call AddLogEntry(LOG_TYPE_ERROR, strFunctionName, strMessage)
                intReturn = BAD_EXIT
                GAM_Plural = intReturn
                Exit Function
            End If
            
            ' Sätter boolElevated till False för att följande körningar ska köras utan förhöjd behörighet
            boolElevated = False
            
            .MoveNext
        Wend
    End With
    
    ' Stäng Recordset
    rst<Namn>.Close
    
    GAM_Plural = intReturn

End Function