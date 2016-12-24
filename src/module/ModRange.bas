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

    If NameExists(strRange) Then
        Range(strRange).Value2 = varValue
    Else
        ModLog.EnableLogging
        ModLog.LogToFile ModSetting.GetLogPath, Error, "Could not set " & varValue & " to range: " & strRange
        ModLog.DisableLogging
    End If

End Sub

Public Function GetRangeValue(strRange As String, varDefault As Variant) As Variant

    If NameExists(strRange) Then
        GetRangeValue = Range(strRange).Value2
    Else
        ModLog.EnableLogging
        ModLog.LogToFile ModSetting.GetLogPath, Error, "Could not get range value from: " & strRange
        ModLog.DisableLogging
        GetRangeValue = varDefault
    End If

End Function

Public Function GetCellAddress(objRange As Range) As String

    GetCellAddress = "=" & "'" & objRange.Parent.Name & "'!" & objRange.Address(External:=False)

End Function

