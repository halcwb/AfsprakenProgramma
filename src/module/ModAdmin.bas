Attribute VB_Name = "ModAdmin"
Option Explicit

Private Enum ConfigTable
    PedCont = 1
    NeoCont = 2
    Parent = 3
End Enum

Private Function Util_CheckPassword() As Boolean
    
    Dim blnValidPw As Boolean
    
    blnValidPw = True
    If Not ModMessage.ShowPasswordBox("Voer admin paswoord in") = ModConst.CONST_PASSWORD Then
        ModMessage.ShowMsgBoxExclam "Deze functie kan alleen met een geldig admin passwoord worden gebruikt"
        blnValidPw = False
    End If
    
    Util_CheckPassword = blnValidPw

End Function

Public Sub Admin_ShowColorPicker()

    Dim frmPicker As FormColorPicker
    
    If Not Util_CheckPassword Then Exit Sub
    
    Set frmPicker = New FormColorPicker
    
    frmPicker.Show
    
End Sub

' ToDo add methods to setup data files and refresh patient data admin jobs

Private Sub Util_SetUpDataDir(ByVal strBedsFilePath As String, arrBeds() As Variant)
    
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

Public Sub Admin_SetUpPedDataDir()
    
    Dim arrBeds() As Variant
    Dim strBedsFilePath As String
        
    If Not Util_CheckPassword Then Exit Sub
    
    arrBeds = ModSetting.GetPedBeds()
    strBedsFilePath = ModSetting.GetPatientDataPath() & ModSetting.CONST_PICU_BEDS
    
    Util_SetUpDataDir strBedsFilePath, arrBeds
    
    ModMessage.ShowMsgBoxInfo "Data bestanden aangemaakt voor afdeling Pediatrie"

End Sub

Public Sub Admin_SetUpNeoDataDir()
    
    Dim arrBeds() As Variant
    Dim strBedsFilePath As String
    
    If Not Util_CheckPassword Then Exit Sub
    
    arrBeds = ModSetting.GetNeoBeds()
    strBedsFilePath = ModSetting.GetPatientDataPath() & ModSetting.CONST_NICU_BEDS
    
    Util_SetUpDataDir strBedsFilePath, arrBeds
    
    ModMessage.ShowMsgBoxInfo "Data bestanden aangemaakt voor afdeling Neonatologie"

End Sub

Public Sub Admin_OpenLogFiles()

    Dim objForm As FormLog
    
    Set objForm = New FormLog
    
    objForm.Show
    
End Sub

Private Sub Util_SelectAdminSheet(objSheet As Worksheet, objRange As Range, strTitle As String)

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

Public Sub Admin_MedContNeoConfig()
    
    Dim frmConfig As FormAdminNeoMedCont
    
    Set frmConfig = New FormAdminNeoMedCont
    
    frmConfig.Show
    
End Sub


Public Sub Admin_ParEntConfig()
    
    Dim frmConfig As FormAdminParent
    
    Set frmConfig = New FormAdminParent
    
    frmConfig.Show
    
End Sub

Public Function Admin_MedContNeoGetVerdunning() As String

    Admin_MedContNeoGetVerdunning = ModRange.GetRangeValue(CONST_MEDCONTVERDUNNING_NEO, vbNullString)

End Function

Public Function Admin_MedContNeoGetCollection() As Collection

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
        objMed.Generic = objTable.Cells(intN, 1)
        objMed.GenericUnit = objTable.Cells(intN, 2)
        objMed.DoseUnit = objTable.Cells(intN, 3)
        objMed.GenericQuantity = objTable.Cells(intN, 4)
        objMed.GenericVolume = objTable.Cells(intN, 5)
        objMed.MinDose = objTable.Cells(intN, 6)
        objMed.MaxDose = objTable.Cells(intN, 7)
        objMed.AbsMaxDose = objTable.Cells(intN, 8)
        objMed.MinConcentration = objTable.Cells(intN, 9)
        objMed.MaxConcentration = objTable.Cells(intN, 10)
        objMed.Solution = objTable.Cells(intN, 11)
        objMed.SolutionRequired = objTable.Cells(intN, 19)
        objMed.DoseAdvice = objTable.Cells(intN, 12)
        objMed.SolutionVolume = objTable.Cells(intN, 13)
        objMed.DripQuantity = objTable.Cells(intN, 14)
        objMed.Product = objTable.Cells(intN, 15)
        objMed.ShelfLife = objTable.Cells(intN, 16)
        objMed.ShelfCondition = objTable.Cells(intN, 17)
        objMed.PreparationText = objTable.Cells(intN, 18)
        objMed.DilutionText = ModRange.GetRangeValue(CONST_MEDCONTVERDUNNING_NEO, vbNullString)
        
        objCol.Add objMed, objMed.Generic
    Next
    
    Set Admin_MedContNeoGetCollection = objCol

End Function

Public Sub Admin_MedContNeoSetCollection(objNeoMedContCol As Collection, ByVal strVerdunning As String)

    Dim objMed As ClassNeoMedCont
    Dim objTable As Range
    
    Dim intR As Integer
    Dim intN As Integer
    
    ModProgress.StartProgress "Neo Continue Medicatie Configuratie"
    
    Set objTable = ModRange.GetRange("Tbl_Admin_NeoMedCont")
    
    intR = objTable.Rows.Count
    
    For intN = 1 To intR
        
        Set objMed = objNeoMedContCol.Item(intN)
        
        objTable.Cells(intN, 1) = objMed.Generic
        objTable.Cells(intN, 2) = objMed.GenericUnit
        objTable.Cells(intN, 3) = objMed.DoseUnit
        objTable.Cells(intN, 4) = objMed.GenericQuantity
        objTable.Cells(intN, 5) = objMed.GenericVolume
        objTable.Cells(intN, 6) = objMed.MinDose
        objTable.Cells(intN, 7) = objMed.MaxDose
        objTable.Cells(intN, 8) = objMed.AbsMaxDose
        objTable.Cells(intN, 9) = objMed.MinConcentration
        objTable.Cells(intN, 10) = objMed.MaxConcentration
        objTable.Cells(intN, 11) = objMed.Solution
        objTable.Cells(intN, 19) = objMed.SolutionRequired
        objTable.Cells(intN, 12) = objMed.DoseAdvice
        objTable.Cells(intN, 13) = objMed.SolutionVolume
        objTable.Cells(intN, 14) = objMed.DripQuantity
        objTable.Cells(intN, 15) = objMed.Product
        objTable.Cells(intN, 16) = objMed.ShelfLife
        objTable.Cells(intN, 17) = objMed.ShelfCondition
        objTable.Cells(intN, 18) = objMed.PreparationText
        
        ModProgress.SetJobPercentage objMed.Generic & "...", intR, intN
        
    Next
    
    ModRange.SetRangeValue CONST_MEDCONTVERDUNNING_NEO, strVerdunning
    
    ModProgress.FinishProgress
    
End Sub

Public Function Admin_ParEntGetCollection() As Collection

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
    
    Set Admin_ParEntGetCollection = objCol

End Function

Public Sub Admin_ParEntSetCollection(objParEntCol As Collection)

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

Private Sub Util_ExportContMed(ByVal strName As String, ByVal strTable As String, ByVal objHeading As Range, ByVal objSrc As Range)
    
    Dim objWbk As Workbook
    Dim strFile As String
    Dim objDst As Range
    Dim shtDst As Worksheet
    Dim intHeading As Integer
    Dim intTableRows As Integer
    Dim intTableColumns As Integer
    Dim varDir As Variant
    
    On Error GoTo ErrorHandler
    
    varDir = ModFile.GetFolderWithDialog()
    
    If CStr(varDir) = vbNullString Then Exit Sub
    
    strFile = Replace(ModDate.FormatDateTimeSeconds(Now()), ":", "_")
    strFile = Replace(strFile, " ", "_")
    strFile = CStr(varDir) & "\" & strName & "_" & strFile & "_.xlsx"
    
    Set objWbk = Workbooks.Add()
    Set shtDst = objWbk.Sheets(1)
    shtDst.Name = strTable
    
    objHeading.Copy
    shtDst.Range("A1").PasteSpecial xlPasteValues
    
    objSrc.Copy
    intHeading = objHeading.Rows.Count + 1
    intTableRows = objSrc.Rows.Count
    intTableColumns = objSrc.Columns.Count
    shtDst.Range("A" & intHeading).PasteSpecial xlPasteValues
    shtDst.Range(Cells(intHeading, 1), Cells(intTableRows + intHeading - 1, intTableColumns)).Name = strTable
    
    Application.DisplayAlerts = False
    objWbk.SaveAs strFile
    Application.DisplayAlerts = True
    
    objWbk.Close
    
    ModMessage.ShowMsgBoxInfo "Configuratie van continue medicatie geexporteerd naar: " & strFile
    
    Exit Sub

ErrorHandler:

    Application.DisplayAlerts = False
    objWbk.Close
    Application.DisplayAlerts = True
    
    ModMessage.ShowMsgBoxError "Kon configuratie voor continue medicatie niet exporteren"

End Sub

Public Sub Admin_ParentImport()
    

End Sub

Public Sub Admin_ParentExport()


End Sub

Public Sub Admin_MedDiscImport()

    Database_ImportConfigMedDisc

End Sub

Public Sub Admin_MedDiscExport()

    Formularium_ExportMedDiscConfig True

End Sub

Public Sub Admin_MedContPedExport()
    
    Dim objHeading As Range
    Dim objSrc As Range
    Dim strName As String
    Dim varDir As Variant
    
    strName = "PedMedCont"
    
    Set objHeading = shtPedTblMedIV.Range("B1:S3")
    Set objSrc = shtPedTblMedIV.Range(CONST_TBL_MEDCONT_PED)

    Util_ExportContMed strName, CONST_TBL_MEDCONT_PED, objHeading, objSrc

End Sub

' ToDo add dilution text
Public Sub Admin_MedContNeoExport()
    
    Dim objHeading As Range
    Dim objSrc As Range
    Dim strName As String
    Dim varDir As Variant
    
    strName = "NeoMedCont"
    
    Set objHeading = shtNeoTblMedIV.Range("B2:T2")
    Set objSrc = shtNeoTblMedIV.Range(CONST_TBL_MEDCONT_NEO)

    Util_ExportContMed strName, CONST_TBL_MEDCONT_NEO, objHeading, objSrc

End Sub

Public Sub Admin_MedContPedImport()

    Dim objConfigWbk As Workbook
    Dim objSrc As Range
    Dim objDst As Range
    Dim lngErr As Long
    Dim strFile As String
    Dim intVersion As Integer
    Dim strMsg As String
    Dim vbAnswer
        
    Dim objMed As ClassNeoMedCont
    
    On Error GoTo HandleError
    
    strMsg = "Kies een Excel bestand met de Pediatrie configuratie voor continue medicatie"
    ModMessage.ShowMsgBoxInfo strMsg
    
    strFile = ModFile.GetFileWithDialog
        
    strMsg = "Dit bestand importeren?" & vbNewLine & strFile
    If ModMessage.ShowMsgBoxYesNo(strMsg) = vbNo Then Exit Sub
       
    Application.DisplayAlerts = False
        
    Set objConfigWbk = Workbooks.Open(strFile, True, True)
    Set objSrc = objConfigWbk.Sheets(CONST_TBL_MEDCONT_PED).Range(CONST_TBL_MEDCONT_PED)
    Set objDst = ModRange.GetRange(CONST_TBL_MEDCONT_PED)
        
    Sheet_CopyRangeFormulaToDst objSrc, objDst
    Database_SavePedConfigMedCont
    
    objConfigWbk.Close
    Application.DisplayAlerts = True
    
    intVersion = Database_GetLatestConfigMedContVersion("Pediatrie")
    strMsg = "De meest recente versie van de pediatrie configuratie van continue medicatie is nu: " & intVersion
    ModMessage.ShowMsgBoxInfo strMsg
    
    Exit Sub
    
HandleError:

    objConfigWbk.Close
    Application.DisplayAlerts = True
    ModLog.LogError Err, "Could not import: " & strFile

End Sub


Public Sub Admin_MedContNeoImport()

'    Dim objConfigWbk As Workbook
'    Dim objSrc As Range
'    Dim objDst As Range
'    Dim lngErr As Long
'    Dim strFile As String
'
'    Dim objMed As ClassNeoMedCont
'
'    strFile = ModFile.GetFileWithDialog
'
'    Dim strMsg As String
'
'    On Error GoTo HandleError
'
'    Application.DisplayAlerts = False
'
'    Set objConfigWbk = Workbooks.Open(strFile, True, True)
'    Set objSrc = objConfigWbk.Sheets(CONST_TBL_MEDCONT_NEO).Range(CONST_TBL_MEDCONT_NEO)
'    Set objDst = ModRange.GetRange(CONST_TBL_MEDCONT_NEO)
'
'    Sheet_CopyRangeFormulaToDst objSrc, objDst
'    Sheet_CopyRangeFormulaToDst objConfigWbk.Sheets(CONST_MEDCONTVERDUNNING_NEO).Range("A1"), ModRange.GetRange(CONST_MEDCONTVERDUNNING_NEO)
'
'    Set m_MedCol = ModAdmin.Admin_MedContNeoGetCollection()
'
'    lbxMedicamenten.Clear
'    For Each objMed In m_MedCol
'        lbxMedicamenten.AddItem objMed.Generic
'    Next
'
'    ClearMedDetails
'
'    objConfigWbk.Close
'    Application.DisplayAlerts = True
'
'    Exit Sub
'
'HandleError:
'
'    objConfigWbk.Close
'    Application.DisplayAlerts = True
'    ModLog.LogError Err, "Could not import: " & strFile

End Sub

Private Sub Util_ConfigImport(ByVal strTable, ByVal intTable As ConfigTable, ByVal strMsg As String)

    Dim objConfigWbk As Workbook
    Dim objSrc As Range
    Dim objDst As Range
    Dim lngErr As Long
    Dim strFile As String
    Dim intVersion As Integer
    Dim vbAnswer
        
    On Error GoTo HandleError
    
    ModMessage.ShowMsgBoxInfo strMsg
    
    strFile = ModFile.GetFileWithDialog
        
    strMsg = "Dit bestand importeren?" & vbNewLine & strFile
    If ModMessage.ShowMsgBoxYesNo(strMsg) = vbNo Then Exit Sub
       
    Application.DisplayAlerts = False
        
    Set objConfigWbk = Workbooks.Open(strFile, True, True)
    Set objSrc = objConfigWbk.Sheets(strTable).Range(strTable)
    Set objDst = ModRange.GetRange(strTable)
        
    Sheet_CopyRangeFormulaToDst objSrc, objDst
    Select Case intTable
    
        Case PedCont
            Database_SavePedConfigMedCont
            intVersion = Database_GetLatestConfigMedContVersion("Pediatrie")
        
        Case NeoCont
            Sheet_CopyRangeFormulaToDst objConfigWbk.Sheets(CONST_MEDCONTVERDUNNING_NEO).Range("A1"), ModRange.GetRange(CONST_MEDCONTVERDUNNING_NEO)
            Database_SaveNeoConfigMedCont
            intVersion = Database_GetLatestConfigMedContVersion("Neonatologie")
        
        Case Parent
            Database_SaveConfigParEnt
            
    End Select
    
    objConfigWbk.Close
    Application.DisplayAlerts = True
    
    strMsg = "De meest recente versie van de configuratie is nu: " & intVersion
    ModMessage.ShowMsgBoxInfo strMsg
    
    Exit Sub
    
HandleError:

    objConfigWbk.Close
    Application.DisplayAlerts = True
    ModLog.LogError Err, "Could not import: " & strFile

End Sub
