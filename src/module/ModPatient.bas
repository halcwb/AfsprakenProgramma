Attribute VB_Name = "ModPatient"
Option Explicit

Private Const constPatNum = "__0_PatNum"
Private Const constBed = "__1_Bed"
Private Const constAN = "__2_AchterNaam"
Private Const constVN = "__3_VoorNaam"
Private Const constGebDatum = "__4_GebDatum"
Private Const constOpnDat = "_Pat_OpnDatum"
Private Const constGewicht = "_Pat__Gewicht"
Private Const constLengte = "_Pat__Lengte"
Private Const constDagen = "_Pat_Dagen"
Private Const constWeken = "_Pat_Weken"
Private Const constGebGew = "_Pat_GebGew"

Public Function GetGewichtFromRange() As Double

    GetGewichtFromRange = Val(ModRange.GetRangeValue(constGewicht, 0)) / 10

End Function

Public Function GetPatientString() As String

    Dim strPat As String
    
    strPat = "Num: " & ModRange.GetRangeValue(constPatNum, "")
    strPat = "Naam: " & ModRange.GetRangeValue(constAN, "")
    strPat = ", " & ModRange.GetRangeValue(constVN, "")
    
    GetPatientString = strPat

End Function

Public Sub OpenPatientLijst(strCaption As String)
    
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
    End With
    
    Set colPats = Nothing
    Set frmPats = Nothing
    
    Exit Sub
    
OpenPatientListError:

    ModMessage.ShowMsgBoxError ModConst.CONST_DEFAULTERROR_MSG
    ModLog.LogError "Cannot OpenPatientLijst(" & strCaption & ")" & ": " & Err.Number
    
End Sub

Private Sub TestPatientLijst()
    Dim strBed As String
    
    MsgBox strBed
    
    'Application.DisplayAlerts = False
    'Workbooks.Open "\\psf\Dropbox\Excel\Afspraken 2016\TestOmgeving\Pelikaan\ICAP\..\ICAP Data\Patienten.xls", True
    'Application.DisplayAlerts = True
    
End Sub

Public Function CreatePatientInfo(strID As String, strBed As String, strAN As String, strVN As String, strBD As String) As ClassPatientInfo

    Dim objInfo As New ClassPatientInfo
    
    objInfo.Id = strID
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

Public Function GetPatientDetails() As ClassPatientDetails

    Dim objPat As New ClassPatientDetails
    
    objPat.PatientID = ModRange.GetRangeValue(constPatNum, vbNullString)
    objPat.AchterNaam = ModRange.GetRangeValue(constAN, vbNullString)
    objPat.VoorNaam = ModRange.GetRangeValue(constVN, vbNullString)
    objPat.GeboorteDatum = ModRange.GetRangeValue(constGebDatum, ModDate.EmptyDate())
    objPat.OpnameDatum = ModRange.GetRangeValue(constOpnDat, ModDate.EmptyDate())
    objPat.Gewicht = ModRange.GetRangeValue(constGewicht, 0) / 10
    objPat.Lengte = ModRange.GetRangeValue(constLengte, 0)
    objPat.GeboorteGewicht = ModRange.GetRangeValue(constGebGew, 0)
    objPat.Weeks = ModRange.GetRangeValue(constWeken, 0)
    objPat.Days = ModRange.GetRangeValue(constDagen, 0)
    
    Set GetPatientDetails = objPat

End Function

Public Sub WritePatientDetails(objPat As ClassPatientDetails)

    ModRange.SetRangeValue constPatNum, objPat.PatientID
    ModRange.SetRangeValue constAN, objPat.AchterNaam
    ModRange.SetRangeValue constVN, objPat.VoorNaam
        
    If Not ModDate.IsEmptyDate(objPat.GeboorteDatum) Then
        ModRange.SetRangeValue constGebDatum, objPat.GeboorteDatum
    End If
    
    If Not ModDate.IsEmptyDate(objPat.OpnameDatum) Then
        ModRange.SetRangeValue constOpnDat, objPat.OpnameDatum
    End If
    
    ModRange.SetRangeValue constGewicht, objPat.Gewicht * 10
    ModRange.SetRangeValue constLengte, objPat.Lengte
    ModRange.SetRangeValue constGebGew, objPat.GeboorteGewicht
    ModRange.SetRangeValue constWeken, objPat.Weeks
    ModRange.SetRangeValue constDagen, objPat.Days

End Sub

Public Sub EnterPatientDetails()

    Dim frmPat As New FormPatient
    Dim objPat As ClassPatientDetails
    
    Set objPat = GetPatientDetails()
    frmPat.SetPatient objPat
    frmPat.Show
    
    WritePatientDetails objPat
    Set frmPat = Nothing

End Sub

Public Sub CopyPatientData()

    Dim intN As Integer
    Dim strRange As String
    Dim varValue As Variant
        
    With shtPatData
        For intN = 2 To .Range("A1").CurrentRegion.Rows.Count
            strRange = .Cells(intN, 1).Value2
            varValue = ModRange.GetRangeValue(strRange, Null)
            
            .Cells(intN, 2).Value2 = varValue
        Next intN
    End With
    
End Sub

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

Public Function ValidWeightKg(dblWeight As Double) As Boolean

    ValidWeightKg = dblWeight > 0.4 And dblWeight < 200

End Function

Public Function ValidLengthCm(dblLen As Double) As Boolean

    ValidLengthCm = dblLen > 30 And dblLen < 250

End Function

Public Function ValidBirthDate(dtmBD As Date) As Boolean

    Dim dtmMin As Date
    
    dtmMin = DateAdd("y", -100, Date)
    
    ValidBirthDate = dtmBD <= Date And dtmBD > dtmMin

End Function

Public Function ValidAdmissionDate(dtmAdm As Date) As Boolean

    Dim dtmMin As Date
    
    dtmMin = DateSerial(2006, 1, 1)
    ValidAdmissionDate = dtmAdm <= Date And dtmAdm > dtmMin

End Function

Private Sub TestValidBirthDate()

    MsgBox ValidBirthDate(Date)
    MsgBox ValidBirthDate(DateAdd("d", 1, Date))
    MsgBox ValidBirthDate(DateSerial(1901, 1, 1))
    MsgBox ValidBirthDate(ModDate.EmptyDate)

End Sub

Public Function ValidDagen(intDay As Integer) As Boolean

    ValidDagen = intDay >= 0 And intDay < 7

End Function

Public Function ValidWeken(intWeek As Integer) As Boolean

    ValidWeken = intWeek > 24 And intWeek < 50

End Function



