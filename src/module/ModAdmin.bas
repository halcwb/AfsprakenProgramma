Attribute VB_Name = "ModAdmin"
Option Explicit

Private Const constNeoMedContTbl = "Tbl_Admin_NeoMedCont"
Private Const constGlobParEntTbl = "Tbl_Admin_ParEnt"
Private Const constPedMedContTbl = "Tbl_Admin_PedMedCont"

Private Const constNeoMedVerdunning = "Var_Neo_MedCont_VerdunningTekst"

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
    
End Sub

Private Sub SelectAdminSheet(objSheet As Worksheet, objRange As Range, strTitle As String)

    Dim objEdit As AllowEditRange
    Dim blnEdit As Boolean
    
    ModMessage.ShowMsgBoxExclam "Nog niet geimplementeerd"
    
    Exit Sub
    
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

    'SelectAdminSheet shtNeoTblMedIV, ModRange.GetRange(constNeoMedContTbl), "NeoMedCont"
    
    Dim frmConfig As FormAdminNeoMedCont
    
    Set frmConfig = New FormAdminNeoMedCont
    
    frmConfig.Show
    
End Sub

Public Sub Admin_TblPedMedCont()

    SelectAdminSheet shtPedTblMedIV, ModRange.GetRange(constPedMedContTbl), "PedMedCont"

End Sub

Public Sub Admin_TblGlobParEnt()

    ' SelectAdminSheet shtGlobTblParEnt, ModRange.GetRange(constGlobParEntTbl), "GlobParEnt"
    
    Dim frmConfig As FormAdminParent
    
    Set frmConfig = New FormAdminParent
    
    frmConfig.Show
    
End Sub

Public Function Admin_GetNeoMedVerdunning() As String

    Admin_GetNeoMedVerdunning = ModRange.GetRangeValue(constNeoMedVerdunning, vbNullString)

End Function

Public Function Admin_GetNeoMedCont() As Collection

    Dim objCol As Collection
    Dim objMed As ClassNeoMedCont
    Dim objTable As Range
    
    Dim intR As Integer
    Dim intN As Integer
    
    Set objCol = New Collection
    Set objTable = ModRange.GetRange("Tbl_Admin_NeoMedCont")
    
    intR = objTable.Rows.Count
    
    For intN = 1 To intR
        Set objMed = New ClassNeoMedCont
        objMed.Name = objTable.Cells(intN, 1)
        objMed.Unit = objTable.Cells(intN, 2)
        objMed.DoseUnit = objTable.Cells(intN, 3)
        objMed.Conc = objTable.Cells(intN, 4)
        objMed.Volume = objTable.Cells(intN, 5)
        objMed.MinDose = objTable.Cells(intN, 6)
        objMed.MaxDose = objTable.Cells(intN, 7)
        objMed.AbsMax = objTable.Cells(intN, 8)
        objMed.MinConc = objTable.Cells(intN, 9)
        objMed.MaxConc = objTable.Cells(intN, 10)
        objMed.OplVlst = objTable.Cells(intN, 11)
        objMed.Advice = objTable.Cells(intN, 12)
        objMed.OplVol = objTable.Cells(intN, 13)
        objMed.Rate = objTable.Cells(intN, 14)
        objMed.Product = objTable.Cells(intN, 15)
        objMed.Houdbaar = objTable.Cells(intN, 16)
        objMed.Bewaar = objTable.Cells(intN, 17)
        objMed.Tekst = objTable.Cells(intN, 18)
        
        objCol.Add objMed, objMed.Name
    Next
    
    Set Admin_GetNeoMedCont = objCol

End Function

Public Sub Admin_SetNeoMedCont(objNeoMedContCol As Collection)

    Dim objMed As ClassNeoMedCont
    Dim objTable As Range
    
    Dim intR As Integer
    Dim intN As Integer
    
    ModProgress.StartProgress "Neo Continue Medicatie Configuratie"
    
    Set objTable = ModRange.GetRange("Tbl_Admin_NeoMedCont")
    
    intR = objTable.Rows.Count
    
    For intN = 1 To intR
        
        Set objMed = objNeoMedContCol.Item(intN)
        
        objTable.Cells(intN, 1) = objMed.Name
        objTable.Cells(intN, 2) = objMed.Unit
        objTable.Cells(intN, 3) = objMed.DoseUnit
        objTable.Cells(intN, 4) = objMed.Conc
        objTable.Cells(intN, 5) = objMed.Volume
        objTable.Cells(intN, 6) = objMed.MinDose
        objTable.Cells(intN, 7) = objMed.MaxDose
        objTable.Cells(intN, 8) = objMed.AbsMax
        objTable.Cells(intN, 9) = objMed.MinConc
        objTable.Cells(intN, 10) = objMed.MaxConc
        objTable.Cells(intN, 11) = objMed.OplVlst
        objTable.Cells(intN, 12) = objMed.Advice
        objTable.Cells(intN, 13) = objMed.OplVol
        objTable.Cells(intN, 14) = objMed.Rate
        objTable.Cells(intN, 15) = objMed.Product
        objTable.Cells(intN, 16) = objMed.Houdbaar
        objTable.Cells(intN, 17) = objMed.Bewaar
        objTable.Cells(intN, 18) = objMed.Tekst
        
        ModProgress.SetJobPercentage objMed.Name & "...", intR, intN
        
    Next
    
    ModProgress.FinishProgress
    
End Sub

Public Function Admin_GetParEnt() As Collection

    Dim objCol As Collection
    Dim objParEnt As ClassParent
    Dim objTable As Range
    
    Dim intR As Integer
    Dim intN As Integer
    
    Set objCol = New Collection
    Set objTable = ModRange.GetRange("Tbl_Admin_ParEnt")
    
    intR = objTable.Rows.Count
    
    For intN = 1 To intR
        Set objParEnt = New ClassParent
        objParEnt.Name = objTable.Cells(intN, 1)
        objParEnt.Energy = objTable.Cells(intN, 2)
        objParEnt.Eiwit = objTable.Cells(intN, 3)
        objParEnt.KH = objTable.Cells(intN, 4)
        objParEnt.Vet = objTable.Cells(intN, 5)
        objParEnt.Na = objTable.Cells(intN, 6)
        objParEnt.K = objTable.Cells(intN, 7)
        objParEnt.Ca = objTable.Cells(intN, 8)
        objParEnt.P = objTable.Cells(intN, 9)
        objParEnt.Mg = objTable.Cells(intN, 10)
        objParEnt.Fe = objTable.Cells(intN, 11)
        objParEnt.VitD = objTable.Cells(intN, 12)
        objParEnt.Cl = objTable.Cells(intN, 13)
        objParEnt.Product = objTable.Cells(intN, 14)
        
        objCol.Add objParEnt, objParEnt.Name
    Next
    
    Set Admin_GetParEnt = objCol

End Function

Public Sub Admin_SetParEnt(objParEntCol As Collection)

    Dim objParEnt As ClassParent
    Dim objTable As Range
    
    Dim intR As Integer
    Dim intN As Integer
    
    ModProgress.StartProgress "Parenteralia Configuratie"
    
    Set objTable = ModRange.GetRange("Tbl_Admin_ParEnt")
    
    intR = objTable.Rows.Count
    
    For intN = 1 To intR
        
        Set objParEnt = objParEntCol.Item(intN)
        objTable.Cells(intN, 2) = objParEnt.Energy
        objTable.Cells(intN, 3) = objParEnt.Eiwit
        objTable.Cells(intN, 4) = objParEnt.KH
        objTable.Cells(intN, 5) = objParEnt.Vet
        objTable.Cells(intN, 6) = objParEnt.Na
        objTable.Cells(intN, 7) = objParEnt.K
        objTable.Cells(intN, 8) = objParEnt.Ca
        objTable.Cells(intN, 9) = objParEnt.P
        objTable.Cells(intN, 10) = objParEnt.Mg
        objTable.Cells(intN, 11) = objParEnt.Fe
        objTable.Cells(intN, 12) = objParEnt.VitD
        objTable.Cells(intN, 13) = objParEnt.Cl
        objTable.Cells(intN, 14) = objParEnt.Product
        
        ModProgress.SetJobPercentage objParEnt.Name & "...", intR, intN
        
    Next
    
    ModProgress.FinishProgress
    
End Sub

