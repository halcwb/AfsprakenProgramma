Attribute VB_Name = "ModPatient"
Option Explicit

Private Const constPatNum As String = "__0_PatNum"
Private Const constAN As String = "__2_AchterNaam"
Private Const constVN As String = "__3_VoorNaam"
Private Const constGebDatum As String = "__4_GebDatum"
Private Const constOpnDat As String = "_Pat_OpnDatum"
Private Const constGewicht As String = "_Pat_Gewicht"
Private Const constLengte As String = "_Pat_Lengte"
Private Const constDagen As String = "_Pat_Dagen"
Private Const constWeken As String = "_Pat_Weken"
Private Const constGebGew As String = "_Pat_GebGew"

Public Function PatientAchterNaam() As String

    PatientAchterNaam = ModRange.GetRangeValue(constAN, vbNullString)

End Function

Public Function PatientVoorNaam() As String

    PatientVoorNaam = ModRange.GetRangeValue(constVN, vbNullString)

End Function

Public Sub Patient_EnterWeight()

    Dim frmGewichtInvoer As FormInvoerNumeriek
    Dim objPatient As ClassPatientDetails
    
    Set frmGewichtInvoer = New FormInvoerNumeriek
    Set objPatient = New ClassPatientDetails
    
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

    Dim frmLengteInvoer As FormInvoerNumeriek
    Dim objPatient As ClassPatientDetails
    
    Set frmLengteInvoer = New FormInvoerNumeriek
    Set objPatient = New ClassPatientDetails
    
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
    
    strPat = "Num: " & ModRange.GetRangeValue(constPatNum, vbNullString)
    strPat = "Naam: " & ModRange.GetRangeValue(constAN, vbNullString)
    strPat = ", " & ModRange.GetRangeValue(constVN, vbNullString)
    
    GetPatientString = strPat

End Function

Public Sub OpenPatientLijst(ByVal strCaption As String)
    
    Dim frmPats As FormPatLijst
    Dim colPats As Collection
    
    On Error GoTo OpenPatientListError
    
    Set colPats = GetPatients()
    Set frmPats = New FormPatLijst
    
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

Public Function CreatePatientInfo(ByVal strID As String, ByVal strBed As String, ByVal strAN As String, ByVal strVN As String, ByVal strBD As String) As ClassPatientInfo

    Dim objInfo As ClassPatientInfo
    
    Set objInfo = New ClassPatientInfo
    objInfo.Id = strID
    objInfo.Bed = strBed
    objInfo.AchterNaam = strAN
    objInfo.VoorNaam = strVN
    objInfo.BirthDate = strBD
    
    Set CreatePatientInfo = objInfo

End Function

Public Function GetPatients() As Collection

    Dim colPatienten As Collection
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
    Set colPatienten = New Collection

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

    Dim objPat As ClassPatientDetails
    Dim dtmBd As Date
    Dim dtmAdm As Date
    
    Set objPat = New ClassPatientDetails
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

Public Sub WritePatientDetails(ByRef objPat As ClassPatientDetails)

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

    Dim frmPat As FormPatient
    Dim objPat As ClassPatientDetails
    
    Set objPat = GetPatientDetails()
    Set frmPat = New FormPatient
    
    frmPat.SetPatient objPat
    frmPat.Show
    
    WritePatientDetails objPat
    Set frmPat = Nothing

End Sub

Public Sub ClearPatientData(ByVal strStartWith As String, ByVal blnShowWarn As Boolean, ByVal blnShowProgress As Boolean)

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

Public Sub PatientClearAll(ByVal blnShowWarn As Boolean, ByVal blnShowProgress As Boolean)
    
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

Public Function ValidWeightKg(ByVal dblWeight As Double) As Boolean

    ValidWeightKg = dblWeight > 0.4 And dblWeight < 200

End Function

Public Function ValidLengthCm(ByVal dblLen As Double) As Boolean

    ValidLengthCm = dblLen > 30 And dblLen < 250

End Function

Public Function ValidBirthDate(ByVal dtmBd As Date, ByVal dtmAdm As Date) As Boolean

    Dim dtmMin As Date
    
    dtmMin = DateAdd("yyyy", -100, DateTime.Date)
    
    ValidBirthDate = dtmBd <= DateTime.Date And dtmBd > dtmMin And dtmBd <= dtmAdm

End Function

Public Function ValidAdmissionDate(ByVal dtmAdm As Date) As Boolean

    Dim dtmMin As Date
    
    dtmMin = DateSerial(2006, 1, 1)
    ValidAdmissionDate = dtmAdm <= DateTime.Date And dtmAdm > dtmMin

End Function

Public Function ValidDagen(ByVal intDay As Integer) As Boolean

    ValidDagen = intDay >= 0 And intDay < 7

End Function

Public Function ValidWeken(ByVal intWeek As Integer) As Boolean

    ValidWeken = intWeek > 24 And intWeek < 50

End Function

Public Function ValidBirthWeight(ByVal intBw As Integer) As Boolean

    ValidBirthWeight = intBw > 400 And intBw < 9999

End Function



