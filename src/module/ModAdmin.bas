Attribute VB_Name = "ModAdmin"
Option Explicit

Private Const constNeoMedContTbl = "Tbl_Admin_NeoMedCont"
Private Const constGlobParentTbl = "Tbl_Admin_ParEnt"
Private Const constPedMedContTbl = "Tbl_Admin_PedMedCont"

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

Private Sub SetUpDataDir(ByVal strBedsFilePath As String, arrBeds() As Variant)
    
    Dim strPath As String
    Dim enmRes As VbMsgBoxResult
    
    strPath = ModSetting.GetPatientDataPath()
    enmRes = ModMessage.ShowMsgBoxYesNo("Alle bestanden in directory " & strPath & " eerst verwijderen?")
    If enmRes = vbYes Then enmRes = ModMessage.ShowMsgBoxYesNo("Zeker weten?")
    
    Application.DisplayAlerts = False
    ModProgress.StartProgress "Opzetten Data Files"

    If enmRes = vbYes Then ModFile.DeleteAllFilesInDir strPath
    ModWorkBook.CreateDataWorkBooks strBedsFilePath, arrBeds, True
    
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

Public Sub ModAdmin_OpenLogFiles()

    Dim objForm As FormLog
    
    Set objForm = New FormLog
    
    objForm.Show
    
    Set objForm = Nothing

End Sub

Private Sub SelectAdminSheet(objSheet As Worksheet, objRange As Range, strTitle As String)

    Dim objEdit As AllowEditRange
    Dim blnEdit As Boolean
    
    blnEdit = False
    For Each objEdit In objSheet.Protection.AllowEditRanges
        If objEdit.Title = strTitle Then
            blnEdit = True
            Exit For
        End If
    Next

    If Not blnEdit Then
        objSheet.Unprotect ModConst.CONST_PASSWORD
        objSheet.Protection.AllowEditRanges.Add Title:=strTitle, Range:=objRange, Password:=ModConst.CONST_PASSWORD
    End If
    
    
    objSheet.Protect ModConst.CONST_PASSWORD
    objSheet.EnableSelection = xlUnlockedCells

    objSheet.Visible = xlSheetVisible
    objSheet.Select
    objRange.Cells(1, 1).Select

End Sub

Public Sub Admin_TblNeoMedCont()

    SelectAdminSheet shtNeoTblMedIV, ModRange.GetRange(constNeoMedContTbl), "NeoMedCont"

End Sub

Public Sub Admin_TblPedMedCont()

    SelectAdminSheet shtPedTblMedIV, ModRange.GetRange(constPedMedContTbl), "PedMedCont"

End Sub

Public Sub Admin_TblGlobParent()

    SelectAdminSheet shtGlobTblParEnt, ModRange.GetRange(constGlobParentTbl), "GlobParEnt"

End Sub

