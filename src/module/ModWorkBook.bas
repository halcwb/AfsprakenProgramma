Attribute VB_Name = "ModWorkBook"
Option Explicit

Private Const constFileReplace As String = "{FILE}"
Private Const constSheetReplace As String = "{SHEET}"
Private Const constNumReplace As String = "{NUM}"
Private Const constPatNum As String = "=IF(ISBLANK('{FILE}{SHEET}'!$B$2),$F${NUM},'{FILE}{SHEET}'!$B$2)"
Private Const constAchterNaam As String = "=IF(ISBLANK(B{NUM}),$F${NUM},'{FILE}{SHEET}'!$B$4)"
Private Const constVoorNaam As String = "=IF(ISBLANK(B{NUM}),$F${NUM},'{FILE}{SHEET}'!$B$5)"
Private Const constGebDat As String = "=IF(ISBLANK(B{NUM}),$F${NUM},'{FILE}{SHEET}'!$B$6)"

Public Sub CreateDataWorkBooks(ByVal strBedsFilePath As String, ByRef arrBeds() As Variant, ByVal strPath As String, ByVal blnShowProgress As Boolean)
    
    Dim objWb As Workbook
    
    Dim shtBeds As Worksheet
    Dim intN As Integer
    Dim intC As Integer
    Dim varBed As Variant
    Dim strPatNum As String
    Dim strAchterNaam As String
    Dim strVoorNaam As String
    Dim strGebDat As String
    Dim strFormula As String
    
    Dim strDataFile As String
    Dim strTextFile As String
    Dim strDataName As String
    
    Dim objDataWb As Workbook
    Dim objTextWb As Workbook
    
    On Error GoTo CreatePatientsWorkBookError
    
    strPatNum = Replace(constPatNum, constSheetReplace, ModSetting.CONST_DATA_SHEET)
    strAchterNaam = Replace(constAchterNaam, constSheetReplace, ModSetting.CONST_DATA_SHEET)
    strVoorNaam = Replace(constVoorNaam, constSheetReplace, ModSetting.CONST_DATA_SHEET)
    strGebDat = Replace(constGebDat, constSheetReplace, ModSetting.CONST_DATA_SHEET)
    
    Set objWb = Workbooks.Add
    
    Set shtBeds = objWb.Sheets(1)
    shtBeds.Name = ModSetting.CONST_BEDS_SHEET
    
    shtBeds.Range("A1").Value2 = "Bed"
    shtBeds.Range("B1").Value2 = "PatientNummer"
    shtBeds.Range("C1").Value2 = "AchterNaam"
    shtBeds.Range("D1").Value2 = "VoorNaam"
    shtBeds.Range("E1").Value2 = "Geboortedatum"
    
    intN = 2
    intC = UBound(arrBeds)
    For Each varBed In arrBeds
                
        strDataFile = ModSetting.GetPatientDataFile(CStr(varBed))
        strTextFile = ModSetting.GetPatientTextFile(CStr(varBed))
        strDataName = ModSetting.GetPatientDataWorkBookName(CStr(varBed))
        
        Set objDataWb = Workbooks.Add
        Set objTextWb = Workbooks.Add
        
        objDataWb.Sheets(1).Name = ModSetting.CONST_DATA_SHEET
        objTextWb.Sheets(1).Name = ModSetting.CONST_DATA_SHEET
                
        If Not ModFile.FileExists(strDataFile) Then
            SaveWorkBookAsShared objDataWb, strDataFile
            ModLog.LogInfo "Created: " & strDataFile
        Else
            ModLog.LogInfo "Already exists: " & strDataFile
        End If
        If Not ModFile.FileExists(strTextFile) Then
            SaveWorkBookAsShared objTextWb, strTextFile
            ModLog.LogInfo "Created: " & strTextFile
        Else
            ModLog.LogInfo "Already exists: " & strTextFile
        End If
        
        objDataWb.Close
        objTextWb.Close
        
        Set objDataWb = Nothing
        Set objTextWb = Nothing
        
        shtBeds.Range("A" & intN).Value2 = varBed
        
        strDataFile = Replace(strDataFile, strDataName, "[" & strDataName & "]")
        
        strFormula = Replace(strPatNum, constFileReplace, strDataFile)
        strFormula = Replace(strFormula, constNumReplace, intN)
        shtBeds.Range("B" & intN).Formula = strFormula
    
        strFormula = Replace(strAchterNaam, constFileReplace, strDataFile)
        strFormula = Replace(strFormula, constNumReplace, intN)
        shtBeds.Range("C" & intN).Formula = strFormula
    
        strFormula = Replace(strVoorNaam, constFileReplace, strDataFile)
        strFormula = Replace(strFormula, constNumReplace, intN)
        shtBeds.Range("D" & intN).Formula = strFormula
    
        strFormula = Replace(strGebDat, constFileReplace, strDataFile)
        strFormula = Replace(strFormula, constNumReplace, intN)
        shtBeds.Range("E" & intN).Formula = strFormula
        
        If blnShowProgress Then
            DoEvents
            ModProgress.SetJobPercentage "Created " & CStr(varBed), intC, intN - 1
        End If
        intN = intN + 1
    
    Next varBed
    
    SaveWorkBookAsShared objWb, strBedsFilePath
    objWb.Close
    
    ModLog.LogInfo "Created: " & strBedsFilePath
    
    Exit Sub
    
CreatePatientsWorkBookError:

    ModMessage.ShowMsgBoxError "Kan patient data file niet aanmaken"
    ModLog.LogError "Cannot create patients workbook: " & Join(Array(strDataFile, strTextFile, strFormula), ", ")

End Sub

Public Sub SaveWorkBookAsShared(ByRef objWorkbook As Workbook, ByVal strFile As String)
    
    If Not objWorkbook.MultiUserEditing Then
        objWorkbook.SaveAs strFile, AccessMode:=xlShared
    End If
     
End Sub

Public Function CopyWorkbookRangeToSheet(ByVal strFile As String, ByVal strBook As String, ByVal strRange As String, ByRef shtTarget As Worksheet, ByVal blnShowProgress As Boolean) As Boolean
    
    Dim strJob As String
    
    On Error GoTo CopyWorkbookRangeToSheetError
    
    ' Guard for non existing file
    If Not ModFile.FileExists(strFile) Then GoTo CopyWorkbookRangeToSheetError
    
    strJob = "Kopieer Data Van File"
    With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
        
        ' Clear the target sheet
        shtTarget.Range("A1").CurrentRegion.Clear
        If blnShowProgress Then ModProgress.SetJobPercentage strJob, 100, 25
        
        ' Open the workbook
        FileSystem.SetAttr strFile, Attributes:=vbNormal
        .Workbooks.Open strFile, True
        ' Make sure the workbook can be shared
        SaveWorkBookAsShared .Workbooks(strBook), strFile
        If blnShowProgress Then ModProgress.SetJobPercentage strJob, 100, 50
        
        ' Copy the range to the target
        .Workbooks(strBook).Sheets(1).Range(strRange).CurrentRegion.Select
        Selection.copy
        shtTarget.Range("A1").PasteSpecial xlPasteValues
        If blnShowProgress Then ModProgress.SetJobPercentage strJob, 100, 75
        
        ' Close the workbook
        .Workbooks(strBook).Close
        If blnShowProgress Then ModProgress.SetJobPercentage strJob, 100, 100
        
        .DisplayAlerts = True
        .ScreenUpdating = True
    End With
        
    CopyWorkbookRangeToSheet = True
        
    Exit Function
    
CopyWorkbookRangeToSheetError:

    If Workbooks.Count = 2 Then Workbooks.Item(2).Close ' To Do Improve by che

    ModLog.LogError "CopyWorkbookRangeToSheet " & strFile & ", " & strBook & ", " & strRange & ", " & shtTarget.Name
    
    Application.DisplayAlerts = True
    CopyWorkbookRangeToSheet = False

End Function
