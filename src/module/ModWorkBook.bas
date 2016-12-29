Attribute VB_Name = "ModWorkBook"
Option Explicit

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
