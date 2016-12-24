Attribute VB_Name = "ModPatient"
Option Explicit

Public Function OpenPatientLijst(strCaption As String) As String
    
    Dim strIndex As String
    Dim objPat As ClassPatient
    Dim frmPats As New FormPatLijst
    Dim colPats As New Collection
    
    Set colPats = GetPatients()
    
    With frmPats
        Application.Cursor = xlWait
        .Caption = ModConst.CONST_APPLICATION_NAME & " " & strCaption
        .lstPatienten.Clear
        For Each objPat In colPats
            .lstPatienten.AddItem objPat
        Next objPat
        Application.Cursor = xlDefault
        .Show
        If .lstPatienten.ListIndex > -1 Then
            Application.Cursor = xlWait
            strIndex = VBA.Left$(.lstPatienten.Text, CONST_BEDNAME_LENGTH)
            Application.Cursor = xlDefault
        End If
        .lstPatienten.Clear
    End With
    
    Set colPats = Nothing
    Set frmPats = Nothing
    
    OpenPatientLijst = strIndex
    
End Function

Public Function GetPatients() As Collection

    Dim colPatienten As New Collection
    Dim intCount As Integer
    Dim strBed As String
    Dim strVn As String
    Dim strAn As String
    Dim strBd As String

    If ModWorkBook.CopyWorkbookRangeToSheet(GetPatientDataPath() + "Patienten.xls", "Patienten.xls", "a1", shtGlobTemp) Then
        With colPatienten
            For intCount = 2 To shtGlobTemp.Range("A1").CurrentRegion.Rows.Count
                With shtGlobTemp
                    strBed = .Cells(intCount, 1).Value2
                    strVn = .Cells(intCount, 2).Value2
                    strAn = .Cells(intCount, 3).Value2
                    strBd = IIf(.Cells(intCount, 4).Value2 <> 0, CDate(.Cells(intCount, 4).Value), vbNullString)
                End With
                .Add strBed & ": " & strVn & " " & strAn & ", " & strBd, strBed
            Next intCount
        End With
    End If

    Set GetPatients = colPatienten
    Set colPatienten = Nothing

End Function

Public Sub EnterPatient()

    Dim frmPatient As New FormPatient
    
    frmPatient.Show
    
    Set frmPatient = Nothing

End Sub

Public Function CopyPatientData() As Boolean

    Dim intN As Integer
    
    On Error Resume Next
    
    With shtPatData
        For intN = 2 To .Range("A1").CurrentRegion.Rows.Count
            .Cells(intN, 4).Formula = Range(.Cells(intN, 1).Value).Formula
        Next intN
    End With
    
    CopyPatientData = True
    
End Function

Public Sub ClearPatient(blnShowWarn As Boolean)
    
    Dim intN As Integer, objResult As VbMsgBoxResult
            
    If blnShowWarn Then
        objResult = ModMessage.ShowMsgBoxYesNo("Afspraken echt verwijderen?")
    Else
        objResult = vbYes
    End If
    
    If objResult = vbYes Then
        Application.Cursor = xlWait
        With shtPatData
            For intN = 2 To .Range("A1").CurrentRegion.Rows.Count
                ModRange.SetRangeValue (.Cells(intN, 1).Value2), .Cells(intN, 3).Value2
            Next intN
        End With
        VerwijderLab
        VerwijderAanvullendeAfspraken
        Application.Cursor = xlDefault
    End If
    
    ModApplication.SetDateToDayFormula
    ModApplication.SetApplicationTitle
    
End Sub

Private Sub Test()

    ClearPatient False

End Sub

Public Function GetPatientDataPath() As String

    Dim strDir As String
    
    strDir = ModSetting.GetDataDir()
    GetPatientDataPath = GetRelativePath(strDir)

End Function

Private Function GetRelativePath(strPath As String) As String

    GetRelativePath = ActiveWorkbook.Path + strPath

End Function

Public Function GetPatientWorkBookName(strBed As String) As String

    GetPatientWorkBookName = "Patient" + strBed + ".xls"

End Function

Public Function GetPatientDataFile(strBed As String) As String

    GetPatientDataFile = GetPatientDataPath + GetPatientWorkBookName(strBed)

End Function



