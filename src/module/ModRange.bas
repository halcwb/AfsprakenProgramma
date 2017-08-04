Attribute VB_Name = "ModRange"
Option Explicit

Private Const constReplEmpty As String = "{EMPTY}"
Private Const constReplRefersTo As String = "{REFERSTO}"
Private Const constRefreshFormula As String = "=IF(ISBLANK({REFERSTO}),{EMPTY},{REFERSTO})"

Private Function GetRow(ByVal sheetName As String, ByVal searchString As String) As Integer

    Dim currentRow As Integer

    WbkAfspraken.Sheets(sheetName).Select
    currentRow = 1
    
    Do While (Strings.LCase(Cells(currentRow, 1).Value) <> Strings.LCase(searchString))
        currentRow = currentRow + 1
    Loop
    
    GetRow = currentRow
    
End Function

Public Sub CopyRangeNamesToRangeNames(arrFrom() As String, arrTo() As String)
    
    Dim intN As Integer
    
    For intN = 0 To UBound(arrFrom)
        ModRange.SetRangeValue arrTo(intN), ModRange.GetRangeValue(arrFrom(intN), vbNullString)
    Next intN
    
End Sub

Public Function CopyTempSheetToNamedRanges(ByVal blnShowProgress As Boolean) As Boolean

    Dim intN As Integer
    Dim intC As Integer
    Dim blnAll As Boolean
    Dim strRange As String
    Dim varValue As Variant
    
    blnAll = True
    With shtGlobTemp
        intC = .Range("A1").CurrentRegion.Rows.Count
        For intN = 2 To intC
            strRange = .Cells(intN, 1).Value2
            varValue = .Cells(intN, 2).Value2
            blnAll = blnAll And ModRange.SetRangeValue(strRange, varValue)
            
            If blnShowProgress Then ModProgress.SetJobPercentage "Kopieer Waarden", intC, intN
        Next intN
    End With
        
    CopyTempSheetToNamedRanges = blnAll

End Function

Public Sub SetNameToRange(ByVal strName As String, objRange As Range)

    ModAssert.AssertTrue objRange.Rows.Count = 1 And objRange.Columns.Count = 1, "Name cannot be set to multi cell", True
    
    If NameExists(strName) Then WbkAfspraken.Names(strName).Delete
    
    If RangeHasName(objRange) Then
        objRange.Name.Name = strName
    Else
        WbkAfspraken.Names.Add Name:=strName, RefersTo:=GetCellAddress(objRange)
    End If

End Sub

Public Function RangeHasName(objRange As Range) As Boolean
    
    On Error GoTo NoName

    RangeHasName = objRange.Name <> vbNullString
    
    Exit Function
    
NoName:
    RangeHasName = False

End Function

Public Function NameExists(ByVal strName As String) As Boolean

'    Dim objName As Name
'
'    For Each objName In WbkAfspraken.Names
'
'        If objName.Name = strName Then
'            NameExists = True
'            Exit Function
'        End If
'
'    Next objName

    On Error GoTo NameExistsError
    
    NameExists = WbkAfspraken.Names(strName).Name = strName
    
    Exit Function
    
NameExistsError:

    NameExists = False

End Function

Private Sub TestNameExists()

    MsgBox NameExists("__0_PatNum")
    MsgBox NameExists("foo")
    
    MsgBox WbkAfspraken.Names("__0_PatNum").RefersToRange.Value2

End Sub

Public Function CreateName(ByVal strName As String, ByVal strGroup As String, ByVal intN As Integer, ByVal intMax As Integer, ByVal blnData As Boolean) As String

    Dim strInt As String
    Dim strResult As String

    If strGroup = vbNullString Then
        strResult = "_" & strName & "_"
    Else
        strResult = IIf(blnData, "_" & strGroup & "_" & strName & "_", strGroup & "_" & strName & "_")
    End If
    
    strInt = CStr(intN)
    Do While Len(strInt) < Len(CStr(intMax))
        strInt = "0" & strInt
    Loop
    
    CreateName = strResult & strInt

End Function

Public Function SetRangeValue(ByVal strRange As String, ByVal varValue As Variant) As Boolean

    Dim blnLog As Boolean
    Dim blnSet As Boolean
    
    On Error GoTo SetRangeValueError

    If NameExists(strRange) Then
        blnSet = True
        WbkAfspraken.Names(strRange).RefersToRange.Value2 = varValue
    Else
        blnLog = ModSetting.GetEnableLogging()
        blnSet = False
        ModLog.EnableLogging
        ModLog.LogToFile ModSetting.GetLogPath(), Error, "Could not set " & varValue & " to range: " & strRange
        If Not blnLog Then ModLog.DisableLogging
    End If
    
    SetRangeValue = blnSet
    Exit Function
    
SetRangeValueError:

    ModLog.LogError "Could not set " & varValue & " to range " & strRange & " Err: " & Err.Number

End Function

Public Sub SetRangeFormula(ByVal strRange As String, ByVal strFormula As String)

    Dim blnLog As Boolean
    
    blnLog = ModSetting.GetEnableLogging()

    If NameExists(strRange) Then
        WbkAfspraken.Names(strRange).RefersToRange.Formula = strFormula
    Else
        ModLog.EnableLogging
        ModLog.LogToFile ModSetting.GetLogPath(), Error, "Could not set " & strFormula & " to range: " & strRange
        If Not blnLog Then ModLog.DisableLogging
    End If

End Sub

Public Function GetRange(ByVal strRange As String) As Range

    Dim blnLog As Boolean
    
    If NameExists(strRange) Then
        Set GetRange = WbkAfspraken.Names(strRange).RefersToRange
    Else
        blnLog = ModSetting.GetEnableLogging()
        ModLog.EnableLogging
        ModLog.LogToFile ModSetting.GetLogPath(), Error, "Could not get range value from: " & strRange
        If Not blnLog Then ModLog.DisableLogging
        
        Set GetRange = Nothing
    End If

End Function

Public Function GetRangeValue(ByVal strRange As String, ByVal varDefault As Variant) As Variant

    Dim blnLog As Boolean
    
    If NameExists(strRange) Then
        GetRangeValue = WbkAfspraken.Names(strRange).RefersToRange.Value2
    Else
        blnLog = ModSetting.GetEnableLogging()
        ModLog.EnableLogging
        ModLog.LogToFile ModSetting.GetLogPath(), Error, "Could not get range value from: " & strRange
        If Not blnLog Then ModLog.DisableLogging
        
        GetRangeValue = varDefault
    End If

End Function

Public Function GetCellAddress(objRange As Range) As String

    Dim strAddress As String
    strAddress = "=" & "'" & objRange.Parent.Name & "'!" & objRange.Address(External:=False)
    GetCellAddress = strAddress

End Function

Public Function IsFormulaValue(ByVal strValue As String) As Boolean

    IsFormulaValue = ModString.StartsWith(strValue, "=")

End Function

Public Function IsDataName(ByVal strName As String) As Boolean

    Dim blnData As Boolean
    
    blnData = ModString.StartsWith(strName, "_")
    blnData = blnData And Not strName = "_xlfn.IFERROR"

    IsDataName = blnData

End Function

Private Sub TestIsDataName()

    MsgBox IsDataName("_Test")

End Sub

Public Function IsPedDataName(ByVal strName As String) As Boolean

    IsPedDataName = ModString.StartsWith(strName, "_Ped")

End Function

Public Function IsNeoDataName(ByVal strName As String) As Boolean

    IsNeoDataName = ModString.StartsWith(strName, "_Neo")

End Function

Public Sub WriteNamesToSheet(shtSheet As Worksheet, ByVal blnShowProgress As Boolean)

    Dim objName As Name
    Dim intN As Integer
    Dim intC As Integer
    Dim blnIsFormula As Boolean
    Dim blnIsData As Boolean
    Dim blnIsNeo As Boolean
    Dim blnIsPed As Boolean
    Dim varValue As Variant
    Dim strEmpty As String
        
    On Error Resume Next
    
    shtSheet.UsedRange.Clear
        
    shtSheet.Cells(1, 1).Value2 = "RefersTo"
    shtSheet.Cells(1, 2).Value2 = "Name"
    shtSheet.Cells(1, 3).Value2 = "ReplaceWith"
    shtSheet.Cells(1, 4).Value2 = "InPatData"
    shtSheet.Cells(1, 5).Value2 = "Value"
    shtSheet.Cells(1, 6).Value2 = "IsFormula"
    shtSheet.Cells(1, 7).Value2 = "IsData"
    shtSheet.Cells(1, 8).Value2 = "IsNeo"
    shtSheet.Cells(1, 9).Value2 = "IsPed"
    
    intN = 2
    intC = WbkAfspraken.Names.Count
    strEmpty = Strings.Chr(34) & Strings.Chr(34)
    For Each objName In WbkAfspraken.Names
        blnIsFormula = IsFormulaValue(Range(objName.Name).Formula)
        blnIsData = IsDataName(objName.Name)
        blnIsNeo = IsNeoDataName(objName.Name)
        blnIsPed = IsPedDataName(objName.Name)
        
        If blnIsFormula Then
            varValue = "F:" & Range(objName.Name).Formula
        Else
            varValue = Range(objName.Name).Value2
        End If
        
        shtSheet.Cells(intN, 1).Value2 = Strings.Replace(objName.RefersTo, "=", vbNullString)
        shtSheet.Cells(intN, 2).Value2 = objName.Name
        shtSheet.Cells(intN, 4).Formula = "=IFERROR(VLOOKUP(B" & intN & ",PatData!$A$2:$A$2000,1,)," & strEmpty & ")<>" & strEmpty
        shtSheet.Cells(intN, 5).Value2 = varValue
        shtSheet.Cells(intN, 6).Value2 = blnIsFormula ' Is Formula
        shtSheet.Cells(intN, 7).Value2 = blnIsData ' Is Data
        shtSheet.Cells(intN, 8).Value2 = blnIsNeo ' Is Neo Data
        shtSheet.Cells(intN, 9).Value2 = blnIsPed ' Is Ped Data
        intN = intN + 1
        
        If blnShowProgress Then ModProgress.SetJobPercentage "Namen Schrijven", intC, intN
        ' If intN = 100 Then Exit For
    Next objName

End Sub

Public Sub WriteNamesToGlobNames()

    Application.ScreenUpdating = False
    ModProgress.StartProgress "Schrijf Namen naar GlobNames Blad"

    WriteNamesToSheet shtGlobNames, True
    
    ModProgress.FinishProgress
    
    Application.ScreenUpdating = True

End Sub

Public Sub ReplaceRangeNames()

    Dim intN As Integer
    Dim intC As Integer
    Dim strOld As String
    Dim strNew As String
    Dim strRefersTo As String
    Dim objRange As Range
    
    ModProgress.StartProgress "Namen Vervangen"
    
    intC = shtGlobNames.Range("A1").CurrentRegion.Count - 1
    For intN = 2 To intC
        strNew = shtGlobNames.Cells(intN, 3).Value2

        If strNew <> vbNullString Then
            strOld = shtGlobNames.Cells(intN, 2).Value2
            
            If ModRange.NameExists(strOld) Then
                WbkAfspraken.Names(strOld).Name = strNew
            Else
                strRefersTo = shtGlobNames.Cells(intN, 1).Value2
                If strRefersTo <> vbNullString Then
                    Set objRange = Range(strRefersTo)
                    SetNameToRange strNew, objRange
                End If
            End If
        End If
        
        ModProgress.SetJobPercentage "Vervang", intC, intN
    
    Next intN
    
    Set objRange = Nothing
    
    ModProgress.FinishProgress
    ModMessage.ShowMsgBoxExclam "Names have been replaced"

End Sub

' Shows the frmNaamGeven to give a range a
' sequential naming of "Name_" + a number
' When runnig this from the visual basic editor
' it works as expected. When running from the ribbon
' menu, the selection in the sheet is not visible.
' But it works as otherwise.
Public Sub GiveNameToRange()

    Dim frmNaamGeven As FormNaamGeven
    
    Set frmNaamGeven = New FormNaamGeven
    frmNaamGeven.Show vbModal
    
    Set frmNaamGeven = Nothing

End Sub

Public Sub RefreshPatientData()

    Dim intN As Integer
    Dim intC As Integer
    Dim strName As String
    Dim strFormula As String
    Dim objName As Name
    
    ModProgress.StartProgress "Ververs Patient Data Blad"
    
    intC = shtPatData.Range("A1").CurrentRegion.Rows.Count
    For intN = 2 To intC
        strName = shtPatData.Cells(intN, 1).Value2
        If NameExists(strName) Then
            Set objName = WbkAfspraken.Names(strName)
            strFormula = Strings.Replace(objName.RefersTo, "=", vbNullString)
            strFormula = Strings.Replace(constRefreshFormula, constReplRefersTo, strFormula)
            strFormula = Strings.Replace(strFormula, constReplEmpty, Chr(34) & Chr(34))
            
            shtPatData.Cells(intN, 2).Formula = strFormula
        End If
        
        ModProgress.SetJobPercentage "Ververs", intC, intN
    Next intN
    
    ModProgress.FinishProgress

End Sub

Public Sub NaamGeven()
    
    Dim frmNaam As FormNaamGeven
    
    Set frmNaam = New FormNaamGeven
    frmNaam.Show
    
    Set frmNaam = Nothing

End Sub


Public Function CollectionFromRange(ByVal strRange As String, ByVal intStart As Integer) As Collection

    Dim colCol As Collection
    Dim intN As Integer
    Dim intC As Integer
    
    On Error GoTo CollectionFromRangeError
    
    Set colCol = New Collection
    intC = Range(strRange).Rows.Count
    For intN = intStart To intC
        colCol.Add Range(strRange).Cells(intN, 1)
    Next intN
    
    Set CollectionFromRange = colCol
    
    Exit Function

CollectionFromRangeError:

    ModLog.LogError "Could not get values from range: " & strRange
    ModMessage.ShowMsgBoxError "Kan waarden niet ophalen"

    Set CollectionFromRange = colCol

End Function


