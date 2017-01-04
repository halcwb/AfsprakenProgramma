Attribute VB_Name = "ModPatient"
Option Explicit

Private Const constPatNum = "__0_PatNum"
Private Const constBed = "__1_Bed"
Private Const constAN = "__2_AchterNaam"
Private Const constVN = "__3_VoorNaam"
Private Const constGebDatum = "__4_GebDatum"
Private Const constOpnDat = "_Pat_OpnDatum"
Private Const constGewicht = "_Pat_Gewicht"
Private Const constLengte = "_Pat_Lengte"
Private Const constDagen = "_Pat_Dagen"
Private Const constWeken = "_Pat_Weken"
Private Const constGebGew = "_Pat_GebGew"

Public Sub Patient_EnterWeight()

    Dim frmGewichtInvoer As New FormInvoerNumeriek
    Dim objPatient As New ClassPatientDetails
    
    With frmGewichtInvoer
        .Caption = "Gewicht invoeren ..."
        .lblParameter.Caption = "Gewicht:"
        .lblEenheid = "kg"
        .txtWaarde = ModRange.GetRangeValue(constGewicht, 0) / 10
        .Show
        If .txtWaarde.Text <> vbNullString Then
            objPatient.Gewicht = .txtWaarde.Text
            If Not IsNull(objPatient.Gewicht) Then
                ModRange.SetRangeValue constGewicht, objPatient.Gewicht * 10
                
            End If
        End If
        .txtWaarde = vbNullString
    End With
    
    PedEntTPN_SelectTPN
    
    Set objPatient = Nothing
    Set frmGewichtInvoer = Nothing

End Sub

Public Sub Patient_EnterLength()

    Dim frmLengteInvoer As New FormInvoerNumeriek
    Dim objPatient As New ClassPatientDetails
    
    With frmLengteInvoer
        .Caption = "Lengte invoeren ..."
        .lblParameter.Caption = "Lengte:"
        .lblEenheid = "cm"
        .txtWaarde = ModRange.GetRangeValue(constLengte, 0)
        .Show
        If .txtWaarde.Text <> vbNullString Then
            objPatient.Lengte = .txtWaarde.Text
            If Not IsNull(objPatient.Lengte) Then
                ModRange.SetRangeValue constLengte, objPatient.Lengte
            End If
        End If
        .txtWaarde = vbNullString
    End With
    
    Set objPatient = Nothing
    Set frmLengteInvoer = Nothing

End Sub

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
    Dim strPatientsName As String
    Dim strPatientsFile As String
    Dim intN As Integer
    Dim strBed As String
    Dim strPN As String
    Dim strVN As String
    Dim strAN As String
    Dim strBD As String

    strPatientsName = ModSetting.GetPatientsFileName()
    strPatientsFile = ModSetting.GetPatientsFilePath()

    If ModWorkBook.CopyWorkbookRangeToSheet(strPatientsFile, strPatientsName, "a1", shtGlobTemp, False) Then
        With colPatienten
            For intN = 2 To shtGlobTemp.Range("A1").CurrentRegion.Rows.Count
                With shtGlobTemp
                    strBed = .Cells(intN, 1).Value2
                    strPN = .Cells(intN, 2).Value2
                    strAN = .Cells(intN, 3).Value2
                    strVN = .Cells(intN, 4).Value2
                    strBD = IIf(.Cells(intN, 5).Value2 > 0, ModString.DateToString(.Cells(intN, 5).Value), vbNullString)
                End With
                .Add CreatePatientInfo(strPN, strBed, strAN, strVN, strBD)
            Next intN
        End With
    End If

    Set GetPatients = colPatienten

End Function

Public Function GetPatientDetails() As ClassPatientDetails

    Dim objPat As New ClassPatientDetails
    Dim dtmBd As Date
    Dim dtmAdm As Date
    
    objPat.PatientID = ModRange.GetRangeValue(constPatNum, vbNullString)
    objPat.AchterNaam = ModRange.GetRangeValue(constAN, vbNullString)
    objPat.VoorNaam = ModRange.GetRangeValue(constVN, vbNullString)
    objPat.Gewicht = ModRange.GetRangeValue(constGewicht, 0) / 10
    objPat.Lengte = ModRange.GetRangeValue(constLengte, 0)
    objPat.GeboorteGewicht = ModRange.GetRangeValue(constGebGew, 0)
    objPat.Weeks = ModRange.GetRangeValue(constWeken, 0)
    objPat.Days = ModRange.GetRangeValue(constDagen, 0)
    
    dtmAdm = ModRange.GetRangeValue(constOpnDat, ModDate.EmptyDate)
    dtmBd = ModRange.GetRangeValue(constGebDatum, ModDate.EmptyDate)
    objPat.SetAdmissionAndBirthDate dtmAdm, dtmBd
    
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

Public Sub ClearPatientData(strStartWith As String, blnShowWarn As Boolean, blnShowProgress As Boolean)

    Dim intN As Integer
    Dim intC As Integer
    Dim strTitle As String
    Dim strRange As String
    Dim varValue As Variant
    Dim strJob As String
    Dim objResult As VbMsgBoxResult
            
    If blnShowWarn Then
        If blnShowProgress Then
            strTitle = FormProgress.Caption
            ModProgress.FinishProgress
        End If
        
        objResult = ModMessage.ShowMsgBoxYesNo("Patient gegevens echt verwijderen?")
    Else
        objResult = vbYes
    End If
    
    If objResult = vbYes Then
        If blnShowProgress And blnShowWarn Then ModProgress.StartProgress strTitle
        
        With shtPatData
            strJob = "Patient gegevens verwijderen ..."
            intC = .Range("A1").CurrentRegion.Rows.Count
            
            For intN = 2 To intC
                strRange = .Cells(intN, 1).Value2
                If strStartWith = vbNullString Then
                    varValue = .Cells(intN, 3).Value2
                    ModRange.SetRangeValue strRange, varValue
                Else
                    If ModString.StartsWith(strRange, strStartWith) Then
                        varValue = .Cells(intN, 3).Value2
                        ModRange.SetRangeValue strRange, varValue
                    End If
                End If
                
                If blnShowProgress Then ModProgress.SetJobPercentage strJob, intC, intN
            Next intN
        End With
        
        ModApplication.SetDateToDayFormula
        ModApplication.SetApplicationTitle
    End If
End Sub

Public Sub PatientClearAll(blnShowWarn As Boolean, blnShowProgress As Boolean)
    
    Dim strTitle As String
    Dim objResult As VbMsgBoxResult
            
    If blnShowWarn Then
        If blnShowProgress Then
            strTitle = FormProgress.Caption
            ModProgress.FinishProgress
        End If
        
        objResult = ModMessage.ShowMsgBoxYesNo("Alle patient gegevens echt verwijderen?")
    Else
        objResult = vbYes
    End If
    
    If objResult = vbYes Then
        ModProgress.StartProgress strTitle
        ClearPatientData vbNullString, False, True
        ModApplication.SetDateToDayFormula
        ModApplication.SetApplicationTitle
    End If
    
End Sub

Private Sub TestClearPatient()
    
    ModProgress.StartProgress "Test Patient Gegevens Verwijderen"
    PatientClearAll False, True
    ModProgress.FinishProgress

End Sub

Public Function ValidWeightKg(dblWeight As Double) As Boolean

    ValidWeightKg = dblWeight > 0.4 And dblWeight < 200

End Function

Public Function ValidLengthCm(dblLen As Double) As Boolean

    ValidLengthCm = dblLen > 30 And dblLen < 250

End Function

Public Function ValidBirthDate(dtmBd As Date, dtmAdm As Date) As Boolean

    Dim dtmMin As Date
    
    dtmMin = DateAdd("yyyy", -100, Date)
    
    ValidBirthDate = dtmBd <= Date And dtmBd > dtmMin And dtmBd <= dtmAdm

End Function

Public Function ValidAdmissionDate(dtmAdm As Date) As Boolean

    Dim dtmMin As Date
    
    dtmMin = DateSerial(2006, 1, 1)
    ValidAdmissionDate = dtmAdm <= Date And dtmAdm > dtmMin

End Function

Public Function ValidDagen(intDay As Integer) As Boolean

    ValidDagen = intDay >= 0 And intDay < 7

End Function

Public Function ValidWeken(intWeek As Integer) As Boolean

    ValidWeken = intWeek > 24 And intWeek < 50

End Function

Public Function ValidBirthWeight(intBw As Integer) As Boolean

    ValidBirthWeight = intBw > 400 And intBw < 9999

End Function



