Attribute VB_Name = "VBAFileDialog"
Option Explicit

' Modulen inneh�ller kod f�r att h�mta s�kv�gar till mapper eller n�r filer ska �ppnas eller sparas.
' Den inneh�ller ett antal standardfilter f�r att t ex �ppna vissa filtyper

' https://docs.microsoft.com/en-us/office/vba/api/office.filedialog

' FileDialog finns i fyra varianter.
' - msoFileDialogFilePicker, h�mtar en s�kv�g till en fil. Kan f�rses med filter
' - msoFileDialogFolderPicker, h�mtar en s�kv�g till en mapp.
' - msoFileDialogOpen, �ppnar en fil som applikationen kan hantera.
' - msoFileDialogSaveAs, sparar en fil i ett format som applikationen kan hantera.

Private Const ACTION_BUTTON_PRESSED = -1
Private Const CANCEL_BUTTON_PRESSED = 0

Public Function VFD_GetFolderPath(Optional strPath As String = vbNullString) As String

    ' L�ter anv�ndaren bl�ddra fram s�kv�gen till en mapp och returnerar s�kv�gen eller vbNullString
    ' Vid upprepade anrop kommer dialogen ih�g mappen om den inte anges explicit
    
    Dim f As FileDialog
    Dim lngReturn As Long
    
    Set f = Application.FileDialog(msoFileDialogFolderPicker)
    
    If strPath <> vbNullString Then
        f.InitialFileName = strPath
    End If
    
    lngReturn = f.Show
    
    If lngReturn = ACTION_BUTTON_PRESSED Then
        VFD_GetFolderPath = f.SelectedItems(1)
        Exit Function
    End If
    
    VFD_GetFolderPath = vbNullString
    
End Function

Public Function VFD_GetTextFileName(Optional strPath As String = vbNullString) As String

    ' �ppnar en dialogruta f�r att h�mta s�kv�gen till en textfil
    
    Dim f As FileDialog
    Dim lngReturn As Long
    Set f = Application.FileDialog(msoFileDialogFilePicker)
    
    If strPath <> vbNullString Then
        f.InitialFileName = strPath
    End If
    
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

Public Function VFD_GetTextFileNameSem(Optional strPath As String = vbNullString) As String

    ' �ppnar en dialogruta f�r att h�mta s�kv�gen till en textfil
    
    Dim f As FileDialog
    Dim lngReturn As Long
    Set f = Application.FileDialog(msoFileDialogFilePicker)
    
    If strPath <> vbNullString Then
        f.InitialFileName = strPath
    End If
    
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

Public Function VFD_GetExcelFileName(Optional strPath As String = vbNullString) As String

    ' �ppnar en dialogruta f�r att spara en Excel-fil
    Dim f As FileDialog
    Dim lngReturn As Long
    Set f = Application.FileDialog(msoFileDialogFilePicker)
    
    If strPath <> vbNullString Then
        f.InitialFileName = strPath
    End If
    
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
