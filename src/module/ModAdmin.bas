Attribute VB_Name = "ModAdmin"
Option Explicit

Private Function CheckPassword() As Boolean
    
    Dim blnValidPw As Boolean
    
    blnValidPw = True
    If Not ModMessage.ShowPasswordBox("Voer admin paswoord in") = ModConst.CONST_PASSWORD Then
        ModMessage.ShowMsgBoxExclam "Deze functie kan alleen met een geldig admin passwoord worden gebruikt"
        blnValidPw = False
    End If
    
    CheckPassword = blnValidPw

End Function


Public Sub ShowColorPicker()

    Dim frmPicker As FormColorPicker
    
    If Not CheckPassword Then Exit Sub
    
    Set frmPicker = New FormColorPicker
    
    frmPicker.Show
    
    Set frmPicker = Nothing

End Sub

' ToDo add methods to setup data files and refresh patient data admin jobs

Private Sub SetUpDataDir(ByVal strBedsFilePath, ByRef arrBeds() As Variant)
    
    Dim strPath As String
    Dim blnDeleteAll As Boolean
    Dim enmRes As VbMsgBoxResult
    
    strPath = ModSetting.GetPatientDataPath()
    enmRes = ModMessage.ShowMsgBoxYesNo("Alle bestanden in directory " & strPath & " eerst verwijderen?")
    If enmRes = vbYes Then enmRes = ModMessage.ShowMsgBoxYesNo("Zeker weten?")
    
    Application.DisplayAlerts = False
    ModProgress.StartProgress "Opzetten Data Files"

    If enmRes = vbYes Then ModFile.DeleteAllFilesInDir strPath
    ModWorkBook.CreateDataWorkBooks strBedsFilePath, arrBeds, strPath, True
    
    ModProgress.FinishProgress
    Application.DisplayAlerts = True
    
End Sub

Public Sub SetUpPedDataDir()
    
    Dim arrBeds() As Variant
    Dim strBedsFilePath As String
        
    If Not CheckPassword Then Exit Sub
    
    arrBeds = ModSetting.GetPedBeds()
    strBedsFilePath = ModSetting.GetPatientDataPath() & ModSetting.CONST_PICU_BEDS
    
    SetUpDataDir strBedsFilePath, arrBeds
    
    ModMessage.ShowMsgBoxInfo "Data bestanden aangemaakt voor afdeling Pediatrie"

End Sub

Public Sub SetUpNeoDataDir()
    
    Dim arrBeds() As Variant
    Dim strBedsFilePath As String
    
    If Not CheckPassword Then Exit Sub
    
    arrBeds = ModSetting.GetNeoBeds()
    strBedsFilePath = ModSetting.GetPatientDataPath() & ModSetting.CONST_NICU_BEDS
    
    SetUpDataDir strBedsFilePath, arrBeds
    
    ModMessage.ShowMsgBoxInfo "Data bestanden aangemaakt voor afdeling Neonatologie"

End Sub


