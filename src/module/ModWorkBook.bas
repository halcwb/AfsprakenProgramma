Attribute VB_Name = "ModWorkBook"
Option Explicit

Private Const constFileReplace = "{File}"
Private Const constSheetReplace = "{Sheet}"
Private Const constQuotesReplace = "{QTS}"
Private Const constNumReplace = "{Num}"
Private Const constPatNum = "=IF('[{File}]{Sheet}'!$B$2={QTS};{QTS};'[{File}]{Sheet}'!$B$2)"
Private Const constAchterNaam = "=IF(B{Num}<>{QTS};'[{File}]{Sheet}'!$B$4;{QTS})"
Private Const constVoorNaam = "=IF(B{Num}<>{QTS};'[{File}]{Sheet}'!$B$5;{QTS})"
Private Const constGebDat = "=IF(B{Num}<>{QTS};'[{File}]{Sheet}'!$B$6;{QTS})"

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
    Dim strQuotes As String
    Dim strDataFile As String
    Dim strTextFile As String
    Dim objDataWb As Workbook
    Dim objTextWb As Workbook
    
    On Error GoTo CreatePatientsWorkBookError
    
    strQuotes = Chr(34) & Chr(34)

    strPatNum = Replace(strPatNum, constSheetReplace, ModSetting.CONST_DATA_SHEET)
    strPatNum = Replace(strPatNum, constQuotesReplace, strQuotes)
    
    strAchterNaam = Replace(strAchterNaam, constSheetReplace, ModSetting.CONST_DATA_SHEET)
    strAchterNaam = Replace(strAchterNaam, constQuotesReplace, strQuotes)
    
    strVoorNaam = Replace(strVoorNaam, constSheetReplace, ModSetting.CONST_DATA_SHEET)
    strVoorNaam = Replace(strVoorNaam, constQuotesReplace, strQuotes)
    
    strGebDat = Replace(strGebDat, constSheetReplace, ModSetting.CONST_DATA_SHEET)
    strGebDat = Replace(strGebDat, constQuotesReplace, strQuotes)
    
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
        
        shtPats.Range("A" & intN).Value2 = varBed
        
        strFormula = Replace(strPatNum, constFileReplace, strDataFile)
        strFormula = Replace(strFormula, constNumReplace, intN - 1)
        shtPats.Range("B" & intN).FormulaLocal = strFormula
    
        strFormula = Replace(strAchterNaam, constFileReplace, strDataFile)
        strFormula = Replace(strFormula, constNumReplace, intN - 1)
        shtPats.Range("C" & intN).FormulaLocal = strFormula
    
        strFormula = Replace(strVoorNaam, constFileReplace, strDataFile)
        strFormula = Replace(strFormula, constNumReplace, intN - 1)
        shtPats.Range("D" & intN).FormulaLocal = strFormula
    
        strFormula = Replace(strGebDat, constFileReplace, strDataFile)
        strFormula = Replace(strFormula, constNumReplace, intN - 1)
        shtPats.Range("E" & intN).FormulaLocal = strFormula
        
        intN = intN + 1
    
    Next varBed
    
    strPatsFile = ModSetting.GetPatientsFilePath
    SaveWorkBookAsShared objWb, strPatsFile
    objWb.Close
    
    Exit Sub
    
CreatePatientsWorkBookError:

    ModMessage.ShowMsgBoxError ModConst.CONST_DEFAULTERROR_MSG
    ModLog.LogError "Cannot create patients workbook: " & Join(Array(strDataFile, strTextFile), ", ")

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
