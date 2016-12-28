Attribute VB_Name = "ModRange"
Option Explicit

Function GetRow(sheetName As String, searchString As String)

    Dim currentRow As Integer

    Sheets(sheetName).Select
    currentRow = 1
    
    Do While (LCase(Cells(currentRow, 1).Value) <> LCase(searchString))
        currentRow = currentRow + 1
    Loop
    
    GetRow = currentRow
    
End Function

Public Sub CopyTempSheetRangeToRange()

    Dim intCount As Integer

    shtPedGuiLab.Unprotect (ModConst.CONST_PASSWORD)
    With shtGlobTemp
        On Error Resume Next
        For intCount = 2 To .Range("A1").CurrentRegion.Rows.Count
            Range(.Cells(intCount, 1).Value).Formula = .Cells(intCount, 2).Formula
        Next intCount
    End With
    shtPedGuiLab.Protect (ModConst.CONST_PASSWORD)

End Sub

Public Sub SetNameToRange(strName As String, objRange As Range)

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

Public Function NameExists(strName As String) As Boolean

    Dim objName As Name
    
    For Each objName In WbkAfspraken.Names
    
        If objName.Name = strName Then
            NameExists = True
            Exit Function
        End If
    
    Next objName
    
    NameExists = False

End Function

Public Function CreateName(ByVal strName As String, ByVal strGroup As String, ByVal intN As Integer, ByVal intMax As Integer) As String

    Dim strInt As String
    Dim strResult As String

    If strGroup = vbNullString Then
        strResult = "_" & strName & "_"
    Else
        strResult = "_" & strGroup & "_" & strName & "_"
    End If
    
    strInt = CStr(intN)
    Do While Len(strInt) < Len(CStr(intMax))
        strInt = "0" & strInt
    Loop
    
    CreateName = strResult & strInt

End Function

Public Sub SetRangeValue(strRange As String, varValue As Variant)

    Dim blnLog As Boolean
    
    blnLog = ModSetting.GetEnableLogging()

    If NameExists(strRange) Then
        Range(strRange).Value2 = varValue
    Else
        ModLog.EnableLogging
        ModLog.LogToFile ModSetting.GetLogPath(), Error, "Could not set " & varValue & " to range: " & strRange
        If Not blnLog Then ModLog.DisableLogging
    End If

End Sub

Public Sub SetRangeFormula(strRange As String, strFormula As String)

    Dim blnLog As Boolean
    
    blnLog = ModSetting.GetEnableLogging()

    If NameExists(strRange) Then
        Range(strRange).FormulaLocal = strFormula
    Else
        ModLog.EnableLogging
        ModLog.LogToFile ModSetting.GetLogPath(), Error, "Could not set " & strFormula & " to range: " & strRange
        If Not blnLog Then ModLog.DisableLogging
    End If

End Sub

Public Function GetRangeValue(strRange As String, varDefault As Variant) As Variant

    Dim blnLog As Boolean
    
    blnLog = ModSetting.GetEnableLogging()

    If NameExists(strRange) Then
        GetRangeValue = Range(strRange).Value2
    Else
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

Public Function IsFormulaValue(strValue As String) As Boolean

    IsFormulaValue = ModString.StartsWith(strValue, "=")

End Function

Public Function IsDataName(strName As String) As Boolean

    IsDataName = ModString.StartsWith(strName, "_") & (Not strName = "_xlfn.IFERROR")

End Function

Public Function IsPedDataName(strName As String) As Boolean

    IsPedDataName = ModString.StartsWith(strName, "_Ped")

End Function

Public Function IsNeoDataName(strName As String) As Boolean

    IsNeoDataName = ModString.StartsWith(strName, "_Neo")

End Function

Public Sub WriteNamesToSheet(shtSheet As Worksheet)

    Dim objName As Name
    Dim intN As Integer
    Dim blnIsFormula As Boolean
    Dim blnIsData As Boolean
    Dim blnIsNeo As Boolean
    Dim blnIsPed As Boolean
    Dim varValue As Variant
    Dim strEmpty As String
        
    On Error Resume Next
    
    shtSheet.Cells.Clear
        
    shtSheet.Cells(1, 1).Value = "RefersTo"
    shtSheet.Cells(1, 2).Value = "Name"
    shtSheet.Cells(1, 3).Value = "ReplaceWith"
    shtSheet.Cells(1, 4).Value = "InPatData"
    shtSheet.Cells(1, 5).Value = "Value"
    shtSheet.Cells(1, 6).Value = "IsFormula"
    shtSheet.Cells(1, 7).Value = "IsData"
    shtSheet.Cells(1, 8).Value = "IsNeo"
    shtSheet.Cells(1, 9).Value = "IsPed"
    
    intN = 2
    strEmpty = Chr(34) & Chr(34)
    For Each objName In WbkAfspraken.Names
        blnIsFormula = IsFormulaValue(Range(objName.Name).FormulaLocal)
        blnIsData = IsDataName(objName.Name)
        blnIsNeo = IsNeoDataName(objName.Name)
        blnIsPed = IsPedDataName(objName.Name)
        
        If blnIsFormula Then
            varValue = "F:" & Range(objName.Name).FormulaLocal
        Else
            varValue = Range(objName.Name).Value2
        End If
        
        shtSheet.Cells(intN, 1).Value = Strings.Replace(objName.RefersTo, "=", "")
        shtSheet.Cells(intN, 2).Value = objName.Name
        shtSheet.Cells(intN, 4).FormulaLocal = "=IFERROR(VLOOKUP(B" & intN & ";PatData!$A$2:$A$1440;1;);" & strEmpty & ")<>" & strEmpty
        shtSheet.Cells(intN, 5).Value = varValue
        shtSheet.Cells(intN, 6).Value = blnIsFormula ' Is Formula
        shtSheet.Cells(intN, 7).Value = blnIsData ' Is Data
        shtSheet.Cells(intN, 8).Value = blnIsNeo ' Is Neo Data
        shtSheet.Cells(intN, 9).Value = blnIsPed ' Is Ped Data
        intN = intN + 1
        ' If intN = 100 Then Exit For
    Next objName

End Sub

Public Sub WriteNamesToGlobNames()

    Application.ScreenUpdating = False
    Application.Cursor = xlWait

    WriteNamesToSheet shtGlobNames
    
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True

End Sub

Public Sub ReplaceRangeNames()

    Dim intN As Integer
    Dim strOld As String
    Dim strNew As String
    
    For intN = 2 To shtGlobNames.Range("A1").CurrentRegion.Rows.Count - 1
        strNew = shtGlobNames.Cells(intN, 3).Value2

        If strNew <> vbNullString Then
            strOld = shtGlobNames.Cells(intN, 2).Value2
            
            WbkAfspraken.Names(strOld).Name = strNew
        End If
    
    Next intN

End Sub

