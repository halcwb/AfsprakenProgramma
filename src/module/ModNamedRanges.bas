Attribute VB_Name = "ModNamedRanges"
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

Public Sub TestNamedRange()

    Dim intN As Integer
    Dim objName As Name
    
    intN = 0
    
    For Each objName In ActiveWorkbook.Names
        intN = intN + 1
        
        If intN > 1 Then Exit For
        
        MsgBox intN & ". " & objName.Name, vbOKOnly
    
    Next objName

End Sub
