Attribute VB_Name = "ModPatient"
Option Explicit

Public Function OpenPatientLijst(strCaption As String) As String
    
    Dim strIndex As String
    Dim objPat As ClassPatientInfo
    Dim frmPats As New FormPatLijst
    Dim colPats As Collection
    
    On Error GoTo OpenPatientListError
    
    Set colPats = GetPatients()
    
    With frmPats
        .Caption = ModConst.CONST_APPLICATION_NAME & " " & strCaption
        .LoadPatients colPats
        .Show
        
        strIndex = .GetSelectedBed()
    End With
    
    Set colPats = Nothing
    Set frmPats = Nothing
    
    OpenPatientLijst = strIndex
    
    Exit Function
    
OpenPatientListError:

    ModMessage.ShowMsgBoxError ModConst.CONST_DEFAULTERROR_MSG
    ModLog.LogError "Cannot OpenPatientLijst(" & strCaption & ")" & ": " & Err.Number
    
End Function

Private Sub TestPatientLijst()
    Dim strBed As String
    
    strBed = OpenPatientLijst("Test")
    MsgBox strBed
    
    'Application.DisplayAlerts = False
    'Workbooks.Open "\\psf\Dropbox\Excel\Afspraken 2016\TestOmgeving\Pelikaan\ICAP\..\ICAP Data\Patienten.xls", True
    'Application.DisplayAlerts = True
    
End Sub

Public Function CreatePatientInfo(strId As String, strBed As String, strAN As String, strVN As String, strBD As String) As ClassPatientInfo

    Dim objInfo As New ClassPatientInfo
    
    objInfo.Id = strId
    objInfo.Bed = strBed
    objInfo.AchterNaam = strAN
    objInfo.VoorNaam = strVN
    objInfo.BirthDate = strBD
    
    Set CreatePatientInfo = objInfo

End Function

Public Function GetPatients() As Collection

    Dim colPatienten As New Collection
    Dim intCount As Integer
    Dim strBed As String
    Dim strVN As String
    Dim strAN As String
    Dim strBD As String

    If ModWorkBook.CopyWorkbookRangeToSheet(GetPatientDataPath() + "Patienten.xls", "Patienten.xls", "a1", shtGlobTemp) Then
        With colPatienten
            For intCount = 2 To shtGlobTemp.Range("A1").CurrentRegion.Rows.Count
                With shtGlobTemp
                    strBed = .Cells(intCount, 1).Value2
                    strVN = .Cells(intCount, 2).Value2
                    strAN = .Cells(intCount, 3).Value2
                    strBD = IIf(.Cells(intCount, 4).Value2 <> 0, ModString.StringToDate(.Cells(intCount, 4).Value), vbNullString)
                End With
                .Add CreatePatientInfo("", strBed, strAN, strVN, strBD)
            Next intCount
        End With
    End If

    Set GetPatients = colPatienten

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
        
        ClearLab
        ClearAfspraken
        
        ModApplication.SetDateToDayFormula
        ModApplication.SetApplicationTitle
    
        Application.Cursor = xlDefault
    End If
    
End Sub

Private Sub TestClearPatient()

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



