Attribute VB_Name = "ModPatient"
Option Explicit

Private Const constPatNum As String = "__0_PatNum"
Private Const constAN As String = "__2_AchterNaam"
Private Const constVN As String = "__3_VoorNaam"
Private Const constGebDatum As String = "__4_GebDatum"
Private Const constOpnDat As String = "_Pat_OpnDatum"
Private Const constGewicht As String = "_Pat_Gewicht"
Private Const constLengte As String = "_Pat_Lengte"
Private Const constGeslacht As String = "_Pat_Geslacht"
Private Const constDagen As String = "_Pat_GestDagen"
Private Const constWeken As String = "_Pat_GestWeken"
Private Const constGebGew As String = "_Pat_GebGew"

Private Const constDateFormatDutch As String = "dd-mmm-jj"
Private Const constDateFormatEnglish As String = "dd-mmm-yy"
Private Const constDateReplace As String = "{DATEFORMAT}"
Private Const constDateFormula As String = vbNullString
Private Const constOpnameDate As String = "Var_Pat_OpnameDat"

Public Function Patient_BirthDate() As Date

    Patient_BirthDate = ModRange.GetRangeValue(constGebDatum, DateTime.Now)

End Function

Public Function PatientHospNum() As String

    PatientHospNum = ModRange.GetRangeValue(constPatNum, vbNullString)

End Function

Public Function PatientAchterNaam() As String

    PatientAchterNaam = ModRange.GetRangeValue(constAN, vbNullString)

End Function

Public Function PatientVoorNaam() As String

    PatientVoorNaam = ModRange.GetRangeValue(constVN, vbNullString)

End Function

Public Sub Patient_EnterWeight()

    Dim frmInvoer As FormInvoerNumeriek
    Dim dblWeight As Double
    
    dblWeight = GetGewichtFromRange()
    Set frmInvoer = New FormInvoerNumeriek
    
    With frmInvoer
        .lblText.Caption = "Voer gewicht in"
        .SetValue vbNullString, "Gewicht:", dblWeight, "kg", "Gewicht"
        
        .Show
        
        If Not .txtWaarde.Value = vbNullString Then
            dblWeight = StringToDouble(.txtWaarde.Value)
            dblWeight = ModString.FixPrecision(dblWeight, 2)
'            dblWeight = dblWeight * 10
            ModRange.SetRangeValue constGewicht, dblWeight
        End If
    End With
    
    ModPedEntTPN.PedEntTPN_SelectStandardTPN
    
End Sub

Public Sub Patient_EnterLength()

    Dim frmInvoer As FormInvoerNumeriek
    Dim dblLength As Double
    
    dblLength = ModRange.GetRangeValue(constLengte, 0)
    Set frmInvoer = New FormInvoerNumeriek
    
    With frmInvoer
        .lblText.Caption = "Voer lengte in"
        .SetValue constLengte, "Lengte:", dblLength, "cm", "Lengte"
        .Show
        
    End With
    
End Sub

Public Function GetGewichtFromRange() As Double

    GetGewichtFromRange = StringToDouble(ModRange.GetRangeValue(constGewicht, 0)) ' / 10

End Function

Public Function GetLengteFromRange() As Double

    GetLengteFromRange = StringToDouble(ModRange.GetRangeValue(constLengte, 0)) ' / 10

End Function

Public Function CalculateBSA() As Double
    
    Dim dblW As Double
    Dim dblH As Double
    
    dblW = GetGewichtFromRange()
    dblH = GetLengteFromRange()

    CalculateBSA = Application.WorksheetFunction.Power(dblW, 0.5378) * Application.WorksheetFunction.Power(dblH, 0.3964) * 0.024265

End Function

Private Sub Test_CalculateBSA()

    MsgBox CalculateBSA()

End Sub

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
    
    Exit Sub
    
OpenPatientListError:

    ModMessage.ShowMsgBoxError "Kan patient lijst niet openen"
    ModLog.LogError "Cannot OpenPatientLijst(" & strCaption & ")" & ": " & Err.Number
    
End Sub

Public Function CreatePatientInfo(ByVal strId As String, ByVal strBed As String, ByVal strAN As String, ByVal strVN As String, ByVal strBD As String) As ClassPatientInfo

    Dim objInfo As ClassPatientInfo
    
    Set objInfo = New ClassPatientInfo
    objInfo.Id = strId
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
    strPatientsFile = ModSetting.GetPatientsFilePath(strPatientsName)
    Set colPatienten = New Collection

    If ModWorkBook.CopyWorkbookRangeToSheet(strPatientsFile, strPatientsName, "A1", shtGlobTemp, False) Then
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

Public Sub GetPatientDetails(objPat As ClassPatientDetails)

    Dim dtmBd As Date
    Dim dtmAdm As Date
    
    objPat.PatientId = ModRange.GetRangeValue(constPatNum, vbNullString)
    objPat.Bed = ModBed.GetBed()
    objPat.AchterNaam = ModRange.GetRangeValue(constAN, vbNullString)
    objPat.VoorNaam = ModRange.GetRangeValue(constVN, vbNullString)
    objPat.Gewicht = ModRange.GetRangeValue(constGewicht, 0) ' / 10
    objPat.Lengte = ModRange.GetRangeValue(constLengte, 0)
    objPat.Geslacht = ModRange.GetRangeValue(constGeslacht, vbNullString)
    objPat.GeboorteGewicht = ModRange.GetRangeValue(constGebGew, 0)
    objPat.Weeks = ModRange.GetRangeValue(constWeken, 0)
    objPat.Days = ModRange.GetRangeValue(constDagen, 0)
    
    dtmAdm = ModRange.GetRangeValue(constOpnDat, ModDate.EmptyDate)
    dtmBd = ModRange.GetRangeValue(constGebDatum, ModDate.EmptyDate)
    objPat.SetAdmissionAndBirthDate dtmAdm, dtmBd
    
End Sub

Public Sub WritePatientDetails(objPat As ClassPatientDetails)

    ModRange.SetRangeValue constPatNum, objPat.PatientId
    ModRange.SetRangeValue constAN, objPat.AchterNaam
    ModRange.SetRangeValue constVN, objPat.VoorNaam
        
    If Not ModDate.IsEmptyDate(objPat.GeboorteDatum) Then
        ModRange.SetRangeValue constGebDatum, objPat.GeboorteDatum
    End If
    
    If Not ModDate.IsEmptyDate(objPat.OpnameDatum) Then
        ModRange.SetRangeValue constOpnDat, objPat.OpnameDatum
    End If
    
    ModRange.SetRangeValue constGewicht, objPat.Gewicht ' * 10
    ModRange.SetRangeValue constLengte, objPat.Lengte
    ModRange.SetRangeValue constGeslacht, objPat.Geslacht
    ModRange.SetRangeValue constGebGew, objPat.GeboorteGewicht
    ModRange.SetRangeValue constWeken, objPat.Weeks
    ModRange.SetRangeValue constDagen, objPat.Days
    
    ModBed.SetBed objPat.Bed

End Sub

Public Sub EnterPatientDetails()

    Dim frmPat As FormPatient
    Dim objPat As ClassPatientDetails
    
    Set objPat = New ClassPatientDetails
    GetPatientDetails objPat
    Set frmPat = New FormPatient
    
    frmPat.SetPatient objPat
    frmPat.Show
    
    WritePatientDetails objPat

End Sub

Public Sub ClearPatientData(ByVal strStartWith As String, ByVal blnShowWarn As Boolean, ByVal blnShowProgress As Boolean)

    Dim intN As Integer
    Dim intC As Integer
    Dim strTitle As String
    Dim strRange As String
    Dim varValue As Variant
    Dim strJob As String
    Dim objResult As VbMsgBoxResult
            
    Dim blnInfB As Boolean
    
    With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    
    blnInfB = (MetaVision_IsNeonatologie() Or ModSetting.IsDevelopmentDir()) And Not (strStartWith = "_Ped" Or strStartWith = "_Glob")
            
    If blnShowWarn Then
        If blnShowProgress Then
            strTitle = FormProgress.Caption
            ModProgress.FinishProgress
        End If
        
        DoEvents
        objResult = ModMessage.ShowMsgBoxYesNo("Patient gegevens echt verwijderen?")
    Else
        strTitle = FormProgress.Caption
        objResult = vbYes
    End If
    
    If objResult = vbYes Then
        If blnShowProgress And blnShowWarn Then
            DoEvents
            ModProgress.StartProgress strTitle
        End If
                
        With shtPatData
            strJob = "Patient gegevens verwijderen"
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
        
        If blnInfB Then
            ModNeoInfB.CopyCurrentInfDataToVar blnShowProgress
            ModNeoInfB.NeoInfB_RemoveVoed
        End If
        
        ModRange.SetRangeValue "Var_Neo_PrintApothNo", 0
        
        ModApplication.SetDateToDayFormula
        ModApplication.SetApplicationTitle
    End If
    
    With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With

End Sub

Public Sub PatientClearPed()

    ModProgress.StartProgress "Verwijder Pediatrie Data"
    ClearPatientData "_Ped", False, True
    ClearPatientData "_Glob", False, True
    ModProgress.FinishProgress
    
    ModApplication.SetDateToDayFormula
    ModApplication.SetApplicationTitle

End Sub

Private Sub Test_PatientClearPed()

    PatientClearPed

End Sub

Public Sub PatientClearNeo()

    ModProgress.StartProgress "Verwijder Neo Data"
    ClearPatientData "_Neo", False, True
    ClearPatientData "_Glob", False, True
    ModProgress.FinishProgress
    
    ModApplication.SetDateToDayFormula
    ModApplication.SetApplicationTitle

End Sub

Private Sub Test_PatientClearNeo()

    PatientClearNeo

End Sub
Public Sub PatientClearAll(ByVal blnShowWarn As Boolean, ByVal blnShowProgress As Boolean)
    
    ClearPatientData vbNullString, blnShowWarn, blnShowProgress
    
    ModApplication.SetDateToDayFormula
    ModApplication.SetApplicationTitle
    
End Sub

Private Sub TestClearPatient()
    
    ModProgress.StartProgress "Test Patient Gegevens Verwijderen"
    PatientClearAll False, True
    ModProgress.FinishProgress

End Sub

Public Function ValidWeightKg(ByVal dblWeight As Double) As Boolean

    ValidWeightKg = dblWeight >= 0.4 And dblWeight < 200

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

    ValidWeken = intWeek >= 21 And intWeek < 50

End Function

Public Function ValidBirthWeight(ByVal intBw As Integer) As Boolean

    ValidBirthWeight = intBw >= 400 And intBw < 9999

End Function



