Attribute VB_Name = "ModPublic"
Option Explicit

Private intCount As Integer

Public Sub MakeShared(objWorkbook As Workbook, strFile As String)
    
    If Not objWorkbook.MultiUserEditing Then
        objWorkbook.SaveAs strFile, AccessMode:=xlShared
    End If
     
End Sub

Public Function CopyWorkbookRangeToSheet(strFile As String, strBook As String, strRange As String, objTarget As Worksheet) As Boolean
    
    Dim dtmVersion As Date
    
    On Error GoTo ErrFileOpenen
    
    With Application
        .DisplayAlerts = False
        
        ' Clear the target sheet
        objTarget.Range("A1").CurrentRegion.Clear
        
        ' Open the workbook
        FileSystem.SetAttr strFile, Attributes:=vbNormal
        .Workbooks.Open strFile, True
        
        ' Make sure the workbook can be shared
        MakeShared Workbooks(strBook), strFile
        
        ' Copy the range to the target
        Range(strRange).CurrentRegion.Select
        Selection.Copy
        objTarget.Range("A1").PasteSpecial xlPasteValues
        
        ' Close the workbook
        Workbooks(strBook).Close
        
        'Get the file version
        dtmVersion = FileSystem.FileDateTime(strFile)
        Range("AfsprakenVersie").Value = dtmVersion
        
        .DisplayAlerts = True
    End With
        
    CopyWorkbookRangeToSheet = True
        
    Exit Function
    
ErrFileOpenen:
    MsgBox prompt:="Kan " & strFile & " nu niet openen, probeer dadelijk nog een keer", _
     Buttons:=vbExclamation, Title:="Informedica 2000"
     Application.DisplayAlerts = True
     
    CopyWorkbookRangeToSheet = False

End Function


Public Function CopyWorkBookRangeToTempSheet(strFile As String, strBook As String, strRange As String) As Boolean
    
    Dim dtmVersion As Date
    
    On Error GoTo ErrFileOpenen
    
    With Application
        .DisplayAlerts = False
        
        ' Clear the target sheet
        shtGlobTemp.Range("A1").CurrentRegion.Clear
        FileSystem.SetAttr strFile, Attributes:=vbNormal
        .Workbooks.Open strFile, True
        
        ' Make sure the workbook can be shared
        MakeShared Workbooks(strBook), strFile
        
        Range(strRange).CurrentRegion.Select
        Selection.Copy
        shtGlobTemp.Range("A1").PasteSpecial xlPasteValues
        Workbooks(strBook).Close
        
        'Get the file version
        dtmVersion = FileSystem.FileDateTime(strFile)
        Range("AfsprakenVersie").Value = dtmVersion
        
        .DisplayAlerts = True
    End With
        
    CopyWorkBookRangeToTempSheet = True
        
    Exit Function
    
ErrFileOpenen:
    MsgBox prompt:="Kan " & strFile & " nu niet openen, probeer dadelijk nog een keer", _
     Buttons:=vbExclamation, Title:="Informedica 2000"
     Application.DisplayAlerts = True

End Function

Public Function SaveBedToFile(strFileName As String, sCONST_WORKBOOKNAME As String, strTekstFile, strTekstBookName) As Boolean

    Dim strSelectie As String
    Dim strAfsprakenTekst As String
    Dim dtmVersion As Date
    Dim strMsg As String
    Dim intAnswer As Integer
    
    If Not Range("BedNummer").Value = 0 Then
        dtmVersion = FileSystem.FileDateTime(strFileName)
        If Not dtmVersion = Range("AfsprakenVersie").Value Then
            strMsg = strMsg & "De afspraken zijn inmiddels gewijzig!" & vbNewLine
            strMsg = strMsg & "Wilt u toch de afspraken opslaan?"
            strMsg = strMsg & "opnieuw in"
            intAnswer = MsgBox(strMsg, vbYesNo)
            
            If intAnswer = vbNo Then Exit Function
        End If
    End If
    
'    On Error GoTo ErrFileSave
    
    Application.DisplayAlerts = False
    
    If SavePatient Then
        strSelectie = "a1:B" + CStr(shtPatData.Range("b1").CurrentRegion.Rows.Count)
        strAfsprakenTekst = "a1:c" + CStr(shtPatDataText.Range("c1").CurrentRegion.Rows.Count)
        FileSystem.SetAttr strFileName, Attributes:=vbNormal
        Application.Workbooks.Open strFileName, True
        With Workbooks(sCONST_WORKBOOKNAME)
            .Sheets("Patienten").Cells.Clear
            shtPatData.Range(strSelectie).Copy
            .Sheets("Patienten").Range("A1").PasteSpecial xlPasteValues
            .Save
            .Close
        End With
        Range("AfsprakenVersie").Value = FileSystem.FileDateTime(strFileName)
        
        FileSystem.SetAttr strTekstFile, Attributes:=vbNormal
        Application.Workbooks.Open strTekstFile, True
        With Workbooks(strTekstBookName)
            .Sheets("AfsprakenTekst").Cells.Clear
            shtPatDataText.Range(strAfsprakenTekst).Copy
            .Sheets("AfsprakenTekst").Range("A1").PasteSpecial xlPasteValues
            .Save
            .Close
        End With
            
    End If
    
    Application.DisplayAlerts = True
        
    SaveBedToFile = True
        
    Exit Function
    
ErrFileSave:
    MsgBox prompt:="Kan " & strFileName & " nu niet opslaan, probeer dadelijk nog een keer", _
     Buttons:=vbExclamation, Title:="Informedica 2000"
     Application.DisplayAlerts = True
     Exit Function

End Function

Public Function SavePatient() As Boolean
    
    Dim intCount As Integer
    
    With shtPatData
        For intCount = 2 To .Range("A1").CurrentRegion.Rows.Count
            On Error Resume Next
            .Cells(intCount, 2).Formula = Range(.Cells(intCount, 1).Value).Formula
        Next intCount
    End With
    
    SavePatient = True

End Function

Public Function GetPatients() As Collection

    Dim colPatienten As New Collection
    Dim bed As String, vn As String, an As String, geb As String
    
    If CopyWorkBookRangeToTempSheet(GetPatientDataPath() + "Patienten.xls", "Patienten.xls", "a1") Then
        With colPatienten
            For intCount = 2 To shtGlobTemp.Range("A1").CurrentRegion.Rows.Count
                With shtGlobTemp
                    bed = .Cells(intCount, 1).Value
                    vn = .Cells(intCount, 2).Value
                    an = .Cells(intCount, 3).Value
                    geb = IIf(.Cells(intCount, 4).Value <> 0, CDate(.Cells(intCount, 4).Value), vbNullString)
                End With
                .Add bed & ": " & vn & " " & an & ", " & geb, bed
            Next intCount
        End With
    End If
    
    Set GetPatients = colPatienten
    Set colPatienten = Nothing

End Function

Public Function DeleteRows(oSheet As Worksheet, sKolom As String) As Worksheet

    On Error Resume Next
    
    With oSheet
        For intCount = 1 To 120
            If .Range(sKolom & intCount).Value = "D" Then
                .Rows(intCount).Delete
                intCount = intCount - 1
            End If
        Next intCount
    End With
    
    Set DeleteRows = oSheet
    Set oSheet = Nothing

End Function
