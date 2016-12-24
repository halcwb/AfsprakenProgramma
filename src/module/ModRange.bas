Attribute VB_Name = "ModRange"
Option Explicit

Sub SetTablePerOs()
    
    Dim intStart As Integer
    Dim intEnd As Integer
    Dim strRange As String
    Dim strStartPO As String
    Dim strEndPO As String
    Dim strNamedRangeTablePerOs As String

    intStart = GetRow("tblPO", strStartPO)
    intEnd = GetRow("tblPO", strEndPO)
    
    strRange = "=tblPO!R" & intStart & "C3:R" & intEnd - 1 & "C12"
    ActiveWorkbook.Names.Add Name:=strNamedRangeTablePerOs, RefersToR1C1:=strRange

End Sub

Sub SetTablePoeder()

    Dim intStart As Integer
    Dim intEnd As Integer
    Dim strRange As String
    Dim strStartPoeder As String
    Dim strEndPoeder As String
    Dim strNamedRangeTablePoeder As String

    intStart = GetRow("tblPO", strStartPoeder)
    intEnd = GetRow("tblPO", strEndPoeder)
    
    strRange = "=tblPO!R" & intStart & "C3:R" & intEnd - 1 & "C12"
    ActiveWorkbook.Names.Add Name:=strNamedRangeTablePoeder, RefersToR1C1:=strRange
    
End Sub

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
    
    If NameExists(strName) Then
        WbkAfspraken.Names(strName).Delete
    End If
    
    If NameExists(objRange.Name) Then
        WbkAfspraken.Names(objRange.Name).Name = strName
    Else
        WbkAfspraken.Names.Add Name:=strName, RefersToR1C1:=objRange.Address
    End If

End Sub

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

Public Function CreateName(strName As String, strGroup As String, intN As Integer, intMax As Integer) As String

    Dim strInt As String

    strName = strGroup & "_" & strName & "_"
    
    strInt = CStr(intN)
    Do While Len(strInt) < Len(CStr(intMax))
        strInt = "0" & strInt
    Loop
    
    CreateName = strName & strInt

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

Private Sub Test()

    MsgBox NameExists("KCl")

End Sub


