Attribute VB_Name = "ModWorkBook"
Option Explicit

Private Const constFileReplace = "{FILE}"
Private Const constSheetReplace = "{SHEET}"
Private Const constNumReplace = "{NUM}"
Private Const constPatNum = "=IF(ISBLANK('[{FILE}]{SHEET}'!$B$2);$F${NUM};'[{FILE}]{SHEET}'!$B$2)"
Private Const constAchterNaam = "=IF(ISBLANK(B{NUM});$F${NUM};'[{FILE}]{SHEET}'!$B$4)"
Private Const constVoorNaam = "=IF(ISBLANK(B{NUM});$F${NUM};'[{FILE}]{SHEET}'!$B$5)"
Private Const constGebDat = "=IF(ISBLANK(B{NUM});$F${NUM};'[{FILE}]{SHEET}'!$B$6)"

Private Sub TestFormulas()

    Dim varBed As Variant
    Dim strPatsFile As String
    Dim intN As Integer
    Dim strPatNum As String
    Dim strAchterNaam As String
    Dim strVoorNaam As String
    Dim strGebDat As String
    Dim strFormula As String
    Dim strDataFile As String
    Dim strTextFile As String
    Dim strDataName As String
        
    strPatNum = Replace(constPatNum, constSheetReplace, ModSetting.CONST_DATA_SHEET)
    strAchterNaam = Replace(constAchterNaam, constSheetReplace, ModSetting.CONST_DATA_SHEET)
    strVoorNaam = Replace(constVoorNaam, constSheetReplace, ModSetting.CONST_DATA_SHEET)
    strGebDat = Replace(constGebDat, constSheetReplace, ModSetting.CONST_DATA_SHEET)
            
    intN = 2
    For Each varBed In ModSetting.GetPedBeds()
                
        strDataFile = ModSetting.GetPatientDataFile(CStr(varBed))
        strTextFile = ModSetting.GetPatientTextFile(CStr(varBed))
        strDataName = ModSetting.GetPatientDataWorkBookName(CStr(varBed))
                        
        strFormula = Replace(strPatNum, constFileReplace, strDataName)
        strFormula = Replace(strFormula, constNumReplace, intN)
        ModLog.LogInfo "Set Formula: " & strFormula
    
        strFormula = Replace(strAchterNaam, constFileReplace, strDataName)
        strFormula = Replace(strFormula, constNumReplace, intN)
        ModLog.LogInfo "Set Formula: " & strFormula
    
        strFormula = Replace(strVoorNaam, constFileReplace, strDataName)
        strFormula = Replace(strFormula, constNumReplace, intN)
        ModLog.LogInfo "Set Formula: " & strFormula
    
        strFormula = Replace(strGebDat, constFileReplace, strDataName)
        strFormula = Replace(strFormula, constNumReplace, intN)
        ModLog.LogInfo "Set Formula: " & strFormula
        
        intN = intN + 1
    
    Next varBed
    
End Sub

Public Sub CreateDataWorkBooks(ByRef arrBeds() As Variant, strPath As String)
    
    Dim objWb As Workbook
    
    Dim strPatsFile As String
    Dim shtPats As Worksheet
    Dim intN As Integer
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
    
    Set shtPats = objWb.Sheets(1)
    shtPats.Name = "Patienten"
    
    shtPats.Range("A1").Value2 = "Bed"
    shtPats.Range("B1").Value2 = "PatientNummer"
    shtPats.Range("C1").Value2 = "AchterNaam"
    shtPats.Range("D1").Value2 = "VoorNaam"
    shtPats.Range("E1").Value2 = "Geboortedatum"
    
    intN = 2
    For Each varBed In arrBeds
                
        strDataFile = ModSetting.GetPatientDataFile(CStr(varBed))
        strTextFile = ModSetting.GetPatientTextFile(CStr(varBed))
        strDataName = ModSetting.GetPatientDataWorkBookName(CStr(varBed))
        
        Set objDataWb = Workbooks.Add
        Set objTextWb = Workbooks.Add
        
        objDataWb.Sheets(1).Name = ModSetting.CONST_DATA_SHEET
        objTextWb.Sheets(1).Name = ModSetting.CONST_DATA_SHEET
                
        SaveWorkBookAsShared objDataWb, strDataFile
        SaveWorkBookAsShared objTextWb, strTextFile
        
        objDataWb.Close
        objTextWb.Close
        
        Set objDataWb = Nothing
        Set objTextWb = Nothing
        
        ModLog.LogInfo "Created: " & strDataFile
        ModLog.LogInfo "Created: " & strTextFile
        
        shtPats.Range("A" & intN).Value2 = varBed
        
        strFormula = Replace(strPatNum, constFileReplace, strDataName)
        strFormula = Replace(strFormula, constNumReplace, intN)
        shtPats.Range("B" & intN).FormulaLocal = strFormula
    
        strFormula = Replace(strAchterNaam, constFileReplace, strDataName)
        strFormula = Replace(strFormula, constNumReplace, intN)
        shtPats.Range("C" & intN).FormulaLocal = strFormula
    
        strFormula = Replace(strVoorNaam, constFileReplace, strDataName)
        strFormula = Replace(strFormula, constNumReplace, intN)
        shtPats.Range("D" & intN).FormulaLocal = strFormula
    
        strFormula = Replace(strGebDat, constFileReplace, strDataName)
        strFormula = Replace(strFormula, constNumReplace, intN)
        shtPats.Range("E" & intN).FormulaLocal = strFormula
        
        intN = intN + 1
    
    Next varBed
    
    strPatsFile = ModSetting.GetPatientsFilePath
    SaveWorkBookAsShared objWb, strPatsFile
    objWb.Close
    
    ModLog.LogInfo "Created: " & strPatsFile
    
    Exit Sub
    
CreatePatientsWorkBookError:

    ModMessage.ShowMsgBoxError ModConst.CONST_DEFAULTERROR_MSG
    ModLog.LogError "Cannot create patients workbook: " & Join(Array(strDataFile, strTextFile, strFormula), ", ")

End Sub

Public Sub SaveWorkBookAsShared(objWorkbook As Workbook, strFile As String)
    
    If Not objWorkbook.MultiUserEditing Then
        objWorkbook.SaveAs strFile, AccessMode:=xlShared
    End If
     
End Sub

Public Function CopyWorkbookRangeToSheet(strFile As String, strBook As String, strRange As String, shtTarget As Worksheet) As Boolean
    
    On Error GoTo ErrFileOpenen
    
    With Application
        .DisplayAlerts = False
        
        ' Clear the target sheet
        shtTarget.Range("A1").CurrentRegion.Clear
        
        ' Open the workbook
        FileSystem.SetAttr strFile, Attributes:=vbNormal
        .Workbooks.Open strFile, True
        ' Make sure the workbook can be shared
        SaveWorkBookAsShared .Workbooks(strBook), strFile
        
        ' Copy the range to the target
        .Workbooks(strBook).Sheets(1).Range(strRange).CurrentRegion.Select
        Selection.Copy
        shtTarget.Range("A1").PasteSpecial xlPasteValues
        
        ' Close the workbook
        .Workbooks(strBook).Close
        
        .DisplayAlerts = True
    End With
        
    CopyWorkbookRangeToSheet = True
        
    Exit Function
    
ErrFileOpenen:

    If Workbooks.Count = 2 Then Workbooks.Item(2).Close

    ModLog.LogError "CopyWorkbookRangeToSheet " & strFile & ", " & strBook & ", " & strRange & ", " & shtTarget.Name
    ModMessage.ShowMsgBoxExclam "Kan " & strFile & " nu niet openen, probeer dadelijk nog een keer"
    
    Application.DisplayAlerts = True
    CopyWorkbookRangeToSheet = False

End Function
