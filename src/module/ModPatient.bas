Attribute VB_Name = "ModPatient"
Option Explicit

Private Const constDateFormatDutch As String = "dd-mmm-jj"
Private Const constDateFormatEnglish As String = "dd-mmm-yy"
Private Const constDateReplace As String = "{DATEFORMAT}"
Private Const constDateFormula As String = vbNullString
Private Const constOpnameDate As String = "Var_Pat_OpnameDat"

Private Const constStandardPrefix As String = "standaard_"

Public Sub Patient_SetHospitalNumber(ByVal strHospNum As String)

    ModRange.SetRangeValue CONST_PATHOSPNUM_RANGE, strHospNum

End Sub

Public Function Patient_BirthDate() As Date

    Patient_BirthDate = ModRange.GetRangeValue(CONST_BIRTHDATE_RANGE, DateTime.Now)

End Function

Public Function Patient_CorrectedAgeInMo() As Double

    Dim dtmBD As Date
    Dim intDays As Integer
    Dim intWeeks As Integer
    Dim dtmCorrBD As Date
    Dim dblAge As Double
    
    dtmBD = Patient_BirthDate()
    intDays = ModRange.GetRangeValue(CONST_GESTDAYS_RANGE, 0)
    intWeeks = ModRange.GetRangeValue(CONST_GESTWEEKS_RANGE, 0)
    intDays = (40 * 7) - (intDays + (intWeeks * 7))
    
    If intDays > 0 Then
        dtmCorrBD = DateAdd("d", dtmBD, intDays)
    Else
        dtmCorrBD = dtmBD
    End If
    
    dblAge = DateDiff("m", dtmCorrBD, Now())
    If dblAge < 0 Then dblAge = 0
    
    Patient_CorrectedAgeInMo = dblAge

End Function

Private Sub Test_Patient_CorrectedAgeInMo()


    ModMessage.ShowMsgBoxInfo Patient_CorrectedAgeInMo

End Sub

Public Function Patient_GetHospitalNumber() As String

    Dim strHospNum As String
    
    strHospNum = ModRange.GetRangeValue(CONST_PATHOSPNUM_RANGE, vbNullString)
    
    Patient_GetHospitalNumber = strHospNum

End Function

Public Function Patient_GetLastName() As String

    Patient_GetLastName = ModRange.GetRangeValue(CONST_LASTNAME_RANGE, vbNullString)

End Function

Public Function Patient_GetFirstName() As String

    Patient_GetFirstName = ModRange.GetRangeValue(CONST_FIRSTNAME_RANGE, vbNullString)

End Function

Public Sub Patient_EnterWeight()

    Dim frmInvoer As FormInvoerNumeriek
    Dim dblWeight As Double
    
    dblWeight = Patient_GetWeight()
    Set frmInvoer = New FormInvoerNumeriek
    
    With frmInvoer
        .lblText.Caption = "Voer gewicht in"
        .SetValue vbNullString, "Gewicht:", dblWeight, "kg", "Gewicht"
        
        .Show
        
        If Not .txtWaarde.Value = vbNullString Then
            dblWeight = StringToDouble(.txtWaarde.Value)
            dblWeight = ModString.FixPrecision(dblWeight, 2)
'            dblWeight = dblWeight * 10
            ModRange.SetRangeValue CONST_WEIGHT_RANGE, dblWeight
        End If
    End With
    
    ModPedEntTPN.PedEntTPN_SelectStandardTPN
    
End Sub

Public Sub Patient_EnterLength()

    Dim frmInvoer As FormInvoerNumeriek
    Dim dblLength As Double
    
    dblLength = ModRange.GetRangeValue(CONST_HEIGHT_RANGE, 0)
    Set frmInvoer = New FormInvoerNumeriek
    
    With frmInvoer
        .lblText.Caption = "Voer lengte in"
        .SetValue CONST_HEIGHT_RANGE, "Lengte:", dblLength, "cm", "Lengte"
        .Show
        
    End With
    
End Sub

Public Function Patient_GetWeight() As Double

    Patient_GetWeight = StringToDouble(ModRange.GetRangeValue(CONST_WEIGHT_RANGE, 0)) ' / 10

End Function

Public Function Patient_GetHeight() As Double

    Patient_GetHeight = StringToDouble(ModRange.GetRangeValue(CONST_HEIGHT_RANGE, 0)) ' / 10

End Function

Public Function CalculateBSA() As Double
    
    Dim dblW As Double
    Dim dblH As Double
    
    dblW = Patient_GetWeight()
    dblH = Patient_GetHeight()

    CalculateBSA = Application.WorksheetFunction.Power(dblW, 0.5378) * Application.WorksheetFunction.Power(dblH, 0.3964) * 0.024265

End Function

Private Sub Test_CalculateBSA()

    MsgBox CalculateBSA()

End Sub

Public Function GetPatientString() As String

    Dim strPat As String
    
    strPat = "Num: " & ModRange.GetRangeValue(CONST_PATHOSPNUM_RANGE, vbNullString)
    strPat = "Naam: " & ModRange.GetRangeValue(CONST_LASTNAME_RANGE, vbNullString)
    strPat = ", " & ModRange.GetRangeValue(CONST_FIRSTNAME_RANGE, vbNullString)
    
    GetPatientString = strPat

End Function

Public Sub Patient_OpenFileList(ByVal strCaption As String)
    
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
    ModLog.LogError Err, "Cannot Patient_OpenFileList(" & strCaption & ")" & ": " & Err.Number
    
End Sub

Public Function Patient_OpenDatabaseList(ByVal strCaption As String) As Boolean
    
    Dim frmPats As FormPatLijst
    Dim colPats As Collection
    Dim colStand As Collection
    Dim blnCancel As Boolean
    
    On Error GoTo OpenPatientListError
    
    Set colPats = GetMetaVisionPatients()
    Set colStand = ModDatabase.Database_GetStandardPatients()
    Set frmPats = New FormPatLijst
    
    With frmPats
        .Caption = ModConst.CONST_APPLICATION_NAME & " " & strCaption
        .SetOnlyAdmittedTrue
        .LoadDbPatients colStand, True
        .LoadDbPatients colPats, False
        .Show
    End With
    
    If Not frmPats Is Nothing Then
        blnCancel = frmPats.GetCancel()
        Set frmPats = Nothing
    End If
    
    Patient_OpenDatabaseList = Not blnCancel
    
    Exit Function
    
OpenPatientListError:

    Set frmPats = Nothing
    Patient_OpenDatabaseList = False
    
End Function

Private Sub Test_Patient_OpenDatabaseList()

    Patient_OpenDatabaseList "Test"
    
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

Private Function GetPatients() As Collection

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

Private Function GetMetaVisionPatients() As Collection

    Dim colPats As Collection
    Dim objPat As ClassPatientDetails
    Dim strDep As String
    
    strDep = ModMetaVision.MetaVision_GetDepartment()
    Set colPats = New Collection
    
    ModMetaVision.MetaVision_GetPatientsForDepartment colPats, strDep
    
    Set GetMetaVisionPatients = colPats

End Function

Private Sub GetPatientDetails(objPat As ClassPatientDetails)

    Dim dtmBD As Date
    Dim dtmAdm As Date
    
    objPat.HospitalNumber = ModRange.GetRangeValue(CONST_PATHOSPNUM_RANGE, vbNullString)
    objPat.Bed = ModBed.Bed_GetBedName()
    objPat.AchterNaam = ModRange.GetRangeValue(CONST_LASTNAME_RANGE, vbNullString)
    objPat.VoorNaam = ModRange.GetRangeValue(CONST_FIRSTNAME_RANGE, vbNullString)
    objPat.Gewicht = ModRange.GetRangeValue(CONST_WEIGHT_RANGE, 0) ' / 10
    objPat.Lengte = ModRange.GetRangeValue(CONST_HEIGHT_RANGE, 0)
    objPat.Geslacht = ModRange.GetRangeValue(CONST_GENDER_RANGE, vbNullString)
    objPat.GeboorteGewicht = ModRange.GetRangeValue(CONST_BIRTHWEIGHT_RANGE, 0)
    objPat.Weeks = ModRange.GetRangeValue(CONST_GESTWEEKS_RANGE, 0)
    objPat.Days = ModRange.GetRangeValue(CONST_GESTDAYS_RANGE, 0)
    
    dtmAdm = ModRange.GetRangeValue(CONST_ADMISSIONDATE_RANGE, ModDate.EmptyDate)
    dtmBD = ModRange.GetRangeValue(CONST_BIRTHDATE_RANGE, ModDate.EmptyDate)
    objPat.SetAdmissionAndBirthDate dtmAdm, dtmBD
    
End Sub

Private Sub WritePatientDetails(objPat As ClassPatientDetails)

    ModRange.SetRangeValue CONST_PATHOSPNUM_RANGE, objPat.HospitalNumber
    ModRange.SetRangeValue CONST_LASTNAME_RANGE, objPat.AchterNaam
    ModRange.SetRangeValue CONST_FIRSTNAME_RANGE, objPat.VoorNaam
        
    If Not ModDate.IsEmptyDate(objPat.GeboorteDatum) Then
        ModRange.SetRangeValue CONST_BIRTHDATE_RANGE, objPat.GeboorteDatum
    End If
    
    If Not ModDate.IsEmptyDate(objPat.OpnameDatum) Then
        ModRange.SetRangeValue CONST_ADMISSIONDATE_RANGE, objPat.OpnameDatum
    End If
    
    ModRange.SetRangeValue CONST_WEIGHT_RANGE, objPat.Gewicht ' * 10
    ModRange.SetRangeValue CONST_HEIGHT_RANGE, objPat.Lengte
    ModRange.SetRangeValue CONST_GENDER_RANGE, objPat.Geslacht
    ModRange.SetRangeValue CONST_BIRTHWEIGHT_RANGE, objPat.GeboorteGewicht
    ModRange.SetRangeValue CONST_GESTWEEKS_RANGE, objPat.Weeks
    ModRange.SetRangeValue CONST_GESTDAYS_RANGE, objPat.Days
    
    ModBed.Bed_SetBed objPat.Bed

End Sub

Public Sub Patient_EnterDetails()

    Dim frmPat As FormPatient
    Dim objPat As ClassPatientDetails
    
    Set objPat = New ClassPatientDetails
    GetPatientDetails objPat
    Set frmPat = New FormPatient
    
    frmPat.SetPatient objPat
    frmPat.Show
    
    If Not frmPat.IsCanceled Then WritePatientDetails objPat

End Sub

Public Sub Patient_ClearData(ByVal strStartWith As String, ByVal blnShowWarn As Boolean, ByVal blnShowProgress As Boolean)

    Dim intN As Integer
    Dim intC As Integer
    Dim strTitle As String
    Dim strRange As String
    Dim varValue As Variant
    Dim strJob As String
    Dim strHospNum As String
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
        strHospNum = Patient_GetHospitalNumber()
        
        If blnShowProgress And blnShowWarn Then
            DoEvents
            ModProgress.StartProgress strTitle
        End If
                
        With shtPatData
            strJob = "Patient gegevens verwijderen"
            intC = .Range("A1").CurrentRegion.Rows.Count
            
            For intN = 2 To intC
                strRange = .Cells(intN, 1).Value2
                If Not ModString.StartsWith(strRange, "_User") Then
                    If strStartWith = vbNullString Then
                        varValue = .Cells(intN, 3).Value2
                        ModRange.SetRangeValue strRange, varValue
                    Else
                        If ModString.StartsWith(strRange, strStartWith) Then
                            varValue = .Cells(intN, 3).Value2
                            ModRange.SetRangeValue strRange, varValue
                        End If
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
        
        ModApplication.App_SetPrescriptionsDate
        ModApplication.App_SetApplicationTitle
    End If
    
    With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
    
    Database_LogAction "Clear patient data", ModUser.User_GetCurrent().Login, strHospNum


End Sub

Public Sub Patient_ClearPedData()

    ModProgress.StartProgress "Verwijder Pediatrie Data"
    Patient_ClearData "_Ped", False, True
    Patient_ClearData "_Glob", False, True
    ModProgress.FinishProgress
    
    ModApplication.App_SetPrescriptionsDate
    ModApplication.App_SetApplicationTitle

End Sub

Private Sub Test_Patient_ClearPedData()

    Patient_ClearPedData

End Sub

Public Sub Patient_ClearNeoData()

    ModProgress.StartProgress "Verwijder Neo Data"
    Patient_ClearData "_Neo", False, True
    Patient_ClearData "_Glob", False, True
    ModProgress.FinishProgress
    
    ModApplication.App_SetPrescriptionsDate
    ModApplication.App_SetApplicationTitle

End Sub

Private Sub Test_Patient_ClearNeoData()

    Patient_ClearNeoData

End Sub

Public Sub Patient_ClearAll(ByVal blnShowWarn As Boolean, ByVal blnShowProgress As Boolean)
    
    Patient_ClearData vbNullString, blnShowWarn, blnShowProgress
    
    ModApplication.App_SetPrescriptionsDate
    ModApplication.App_SetApplicationTitle
        
End Sub

Private Sub TestClearPatient()
    
    ModProgress.StartProgress "Test Patient Gegevens Verwijderen"
    Patient_ClearAll False, True
    ModProgress.FinishProgress

End Sub

Public Function ValidWeightKg(ByVal dblWeight As Double) As Boolean

    ValidWeightKg = dblWeight >= 0.4 And dblWeight < 200

End Function

Public Function ValidLengthCm(ByVal dblLen As Double) As Boolean

    ValidLengthCm = dblLen > 30 And dblLen < 250

End Function

Public Function ValidBirthDate(ByVal dtmBD As Date, ByVal dtmAdm As Date) As Boolean

    Dim dtmMin As Date
    
    dtmMin = DateAdd("yyyy", -100, DateTime.Date)
    
    ValidBirthDate = dtmBD <= DateTime.Date And dtmBD > dtmMin And dtmBD <= dtmAdm

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

Public Sub Patient_SavePatient()

    Dim strBed As String
    Dim strNew As String
    
    Dim strPrompt As String
    Dim strAction As String
    Dim strParams() As Variant
    
    Dim varReply As VbMsgBoxResult
    
    Dim blnNeo As Boolean

    On Error GoTo ErrorHandler
    
    MedDisc_SortTableMedDisc
    
    If ModString.StringIsZeroOrEmpty(ModPatient.Patient_GetHospitalNumber()) Then
        ModMessage.ShowMsgBoxExclam "Kan patient zonder ziekenhuis nummer niet opslaan!"
        
        varReply = ModMessage.ShowMsgBoxYesNo("Patient als standaard patient opslaan?")
        If varReply = vbYes Then
            CreateStandardPatient
            ' If still no hosp number, exit
            If ModString.StringIsZeroOrEmpty(ModPatient.Patient_GetHospitalNumber()) Then Exit Sub
        Else
            Exit Sub
        End If
    End If
    
    blnNeo = MetaVision_IsNeonatologie()
    
    strAction = "Patient_SavePatient"
    strParams = Array()
    LogActionStart strAction, strParams
    
    ModProgress.StartProgress "Patient " & ModPatient.GetPatientString & " opslaan"
    
    If blnNeo Then
        ModNeoInfB.NeoInfB_SelectInfB False, False ' Make sure that the Infuusbrief Actueel is selected
        ModNeoInfB.CopyCurrentInfVarToData True    ' Make sure that neo data is updated with latest current infuusbrief
    End If

    If SavePatientToDatabase(True) Then
        ModProgress.FinishProgress
        ModMessage.ShowMsgBoxInfo "Patient is opgeslagen"
    Else
        ModProgress.FinishProgress
        ModMessage.ShowMsgBoxExclam "Patient werd niet opgeslagen"
    End If

    LogActionEnd strAction
    
    Exit Sub
    
ErrorHandler:

    ModProgress.FinishProgress
    
    ModMessage.ShowMsgBoxError "Kan patient niet opslaan in de database"
    ModLog.LogError Err, strAction

End Sub

Private Sub CreateStandardPatient()

    Dim strName As String
    Dim strHospNum As String
    Dim intHospNum As Integer
    
    strName = ModMessage.ShowInputBox("Geef een naam op voor de standaard patient", vbNullString)
    
    If Not strName = vbNullString Then
        strHospNum = ModDatabase.Database_GetLastStandardPatientHospNum()
        strHospNum = Replace(strHospNum, constStandardPrefix, vbNullString)
        If strHospNum = vbNullString Then
            intHospNum = 1
        Else
            intHospNum = CInt(strHospNum) + 1
        End If
        
        If intHospNum > 999 Then
            ModMessage.ShowMsgBoxExclam "Maximaal aantal van 999 standaard patienten is bereikt"
        Else
            strHospNum = constStandardPrefix & intHospNum
            strHospNum = IIf(intHospNum < 100, constStandardPrefix & "0" & intHospNum, strHospNum)
            strHospNum = IIf(intHospNum < 10, constStandardPrefix & "00" & intHospNum, strHospNum)
            
            ModPatient.Patient_SetHospitalNumber strHospNum
            ModRange.SetRangeValue CONST_LASTNAME_RANGE, "Patient"
            ModRange.SetRangeValue CONST_FIRSTNAME_RANGE, strName
            
        End If
        
    End If

End Sub

Private Function GetVersionWarningMsg(ByVal intCurrent As Integer, ByVal intLatest As Integer) As String

    Dim strMsg As String

    strMsg = strMsg & "De afspraken zijn inmiddels gewijzig!" & vbNewLine
    strMsg = strMsg & vbNewLine
    strMsg = strMsg & "De huidige versie is: " & intCurrent & vbNewLine
    strMsg = strMsg & "De laatst opgeslagen versie is: " & intLatest & vbNewLine
    strMsg = strMsg & vbNewLine
    strMsg = strMsg & "Wilt u toch de afspraken opslaan?" & vbNewLine
    strMsg = strMsg & vbNewLine
    strMsg = strMsg & "DEZE VERSIE WORDT DAN DE LAATSTE (ACTUELE) VERSIE!!"

    GetVersionWarningMsg = strMsg

End Function

Private Function GetDateWarningMsg(ByVal strCurrent, ByVal strLatest) As String

    Dim strMsg As String

    strMsg = strMsg & "De afspraken zijn inmiddels gewijzig!" & vbNewLine
    strMsg = strMsg & vbNewLine
    strMsg = strMsg & ModDate.FormatDateHoursMinutes(CDate(strCurrent)) & " (huidige versie) " & vbNewLine
    strMsg = strMsg & ModDate.FormatDateHoursMinutes(CDate(strLatest)) & " (laatst opgeslagen versie)" & vbNewLine
    strMsg = strMsg & vbNewLine
    strMsg = strMsg & "Wilt u toch de afspraken opslaan?" & vbNewLine
    strMsg = strMsg & vbNewLine
    strMsg = strMsg & "U OVERSCHRIJFT DAN RECENTER OPGESLAGEN AFSPRAKEN!!"

    GetDateWarningMsg = strMsg

End Function

Private Sub Test_GetDateWarningMsg()

    ModMessage.ShowMsgBoxYesNo GetDateWarningMsg(FormatDateTimeSeconds(Now()), FormatDateTimeSeconds(Now()))

End Sub

Private Function SavePatientToDatabase(ByVal blnShowProgress As Boolean) As Boolean

    Dim intLatest As Integer
    Dim intCurrent As Integer
    Dim objUser As ClassUser
    Dim strHospNum As String
    Dim strMsg As String
    Dim intC As Integer
    Dim intR As Integer
    Dim varReply As VbMsgBoxResult
        
    Dim strProg As String
        
    On Error GoTo SaveBedToDatabaseError
    
    ' Guard for invalid bed name
    ' If Not IsValidBed(strBed) Then GoTo SaveBedToDatabaseError
    
    intCurrent = ModBed.Bed_PrescriptionsVersionGet()
    strHospNum = Patient_GetHospitalNumber()
    intLatest = ModDatabase.Database_GetLatestPrescriptionVersion(strHospNum)
            
    If Not intLatest = 0 And Not intLatest = intCurrent Then
        If blnShowProgress Then
            strProg = FormProgress.Caption
            ModProgress.FinishProgress
        End If
    
        strMsg = GetVersionWarningMsg(intCurrent, intLatest)
        varReply = ModMessage.ShowMsgBoxYesNo(strMsg)
        
        If blnShowProgress Then ModProgress.StartProgress strProg

        If varReply = vbNo Then Exit Function
    End If
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    ModDatabase.Database_SavePatient
    ModDatabase.Database_SavePrescriber
    
    Set objUser = ModUser.User_GetCurrent()
    
    ModDatabase.Database_SaveData strHospNum, objUser.Login, shtPatData.Range("A1").CurrentRegion, shtPatText.Range("A1").CurrentRegion, blnShowProgress
    ModBed.Bed_PrescriptionsVersionSet Database_GetLatestPrescriptionVersion(strHospNum)
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
        
    SavePatientToDatabase = True
        
    Exit Function
    
SaveBedToDatabaseError:
    
    ModMessage.ShowMsgBoxError "Kan patient niet opslaan"
    ModLog.LogError Err, "Could not save patient to database with: "

    Application.DisplayAlerts = True
    SavePatientToDatabase = False

End Function

Public Sub Patient_OpenPatient()

    ModProgress.StartProgress "Open Patient"
    OpenPatient False, True
    ModProgress.FinishProgress

End Sub

Public Sub Patient_OpenPatientAndAsk()

    ModProgress.StartProgress "Open Patient"
    OpenPatient True, True
    ModProgress.FinishProgress

End Sub

Private Sub OpenPatient(ByVal blnAsk As Boolean, ByVal blnShowProgress As Boolean)
    
    Dim strTitle As String
    Dim strAction As String
    Dim strParams() As Variant
    Dim strFileName As String
    Dim strBookName As String
    Dim strRange As String
    Dim blnAll As Boolean
    Dim blnNeo As Boolean
    Dim strHospNum As String
    Dim strStandard As String
    Dim intVersion As Integer
    
    Dim objPat As ClassPatientDetails
    
    On Error GoTo ErrorHandler
    
    Set objPat = New ClassPatientDetails
    ModMetaVision.MetaVision_GetPatientDetails objPat, ModMetaVision.MetaVision_GetCurrentPatientID, vbNullString
    blnNeo = MetaVision_IsNeonatologie()
    
    strHospNum = Patient_GetHospitalNumber()
    strHospNum = IIf(strHospNum = vbNullString, objPat.HospitalNumber, strHospNum)
    intVersion = ModBed.Bed_PrescriptionsVersionGet()
    ModBed.Bed_PrescriptionsVersionSet 0
    If blnAsk Then
        If blnShowProgress Then
            strTitle = FormProgress.Caption
            ModProgress.FinishProgress
        End If
    
        If Patient_OpenDatabaseList("Selecteer een patient") Then
            If Patient_GetHospitalNumber() = vbNullString Then  ' No patient was selected
                Patient_SetHospitalNumber strHospNum            ' Put back the old hospital number
                Bed_PrescriptionsVersionSet intVersion          ' Put back current version
                Exit Sub                                        ' And exit sub
            Else
                If Patient_IsStandard(Patient_GetHospitalNumber()) Then
                    strStandard = Patient_GetHospitalNumber()
                    Patient_SetHospitalNumber strHospNum
                    strHospNum = vbNullString
                Else
                    strHospNum = Patient_GetHospitalNumber()
                    intVersion = ModBed.Bed_PrescriptionsVersionGet()
                    
                    If blnShowProgress Then ModProgress.StartProgress strTitle
                End If
            End If
        Else                                                ' Patientlist has been canceled so do nothing
            Patient_SetHospitalNumber strHospNum            ' Put back the old hospital number
            Bed_PrescriptionsVersionSet intVersion          ' Put back current version
            
            If blnShowProgress Then ModProgress.FinishProgress
            Exit Sub                                        ' And exit sub
        End If
    End If
    
    If Not strHospNum = vbNullString Then
        Patient_ClearAll False, True
        GetPatientDataFromDatabase strHospNum, ModDatabase.Database_GetLatestPrescriptionVersion(strHospNum)
    End If
    
    If Not strStandard = vbNullString Then
        GetPatientDataFromDatabase strStandard, ModDatabase.Database_GetLatestPrescriptionVersion(strStandard)
    End If
        
    ModDatabase.Database_LogAction "Open Patient", ModUser.User_GetCurrent().Login, strHospNum
    
    Exit Sub

ErrorHandler:

    ModMessage.ShowMsgBoxError "Kan bed " & strHospNum & " niet openenen"
    ModLog.LogError Err, Err.Description
    
End Sub

Public Function Patient_IsStandard(ByVal strHospNum) As Boolean

    Patient_IsStandard = ModString.StartsWith(strHospNum, constStandardPrefix)

End Function

Private Sub Test_OpenPatient()
    
    ModProgress.StartProgress "Testing select patient"
    OpenPatient True, True
    ModProgress.FinishProgress

End Sub

Private Sub GetPatientDataFromDatabase(ByVal strHospNum As String, Optional ByVal intVersion As Integer = 0)
    
    On Error GoTo GetPatientDataFromDatabaseError

    Dim strTitle As String
    Dim strAction As String
    Dim strParams() As Variant
    Dim blnNeo As Boolean
    
    ModProgress.StartProgress "Patient data ophalen voor " & strHospNum
    
    blnNeo = MetaVision_IsNeonatologie()
    
    strAction = "ModPatient.GetPatientDataFromDatabase"
    
    ModLog.LogActionStart strAction, strParams
            
    If blnNeo Then
        ModPatient.Patient_ClearNeoData
    Else
        ModPatient.Patient_ClearPedData
    End If
    
    If intVersion = 0 Then
        ModDatabase.Database_GetPatientData strHospNum
    Else
        ModDatabase.Database_GetPatientDataForVersion strHospNum, intVersion
    End If
    If Not Patient_IsStandard(strHospNum) Then ModRange.SetRangeValue CONST_PATHOSPNUM_RANGE, strHospNum 'Have to set hospitalnumber if leading zero got lost from transfer from database
    
    If blnNeo Then ModNeoInfB.CopyCurrentInfDataToVar True ' Make sure that infuusbrief data is updated
            
    ModApplication.App_SetApplicationTitle

    If Not Patient_IsStandard(strHospNum) Then ModMetaVision.MetaVision_SyncLab
    ModSheet.SelectPedOrNeoStartSheet True
    
    ModProgress.FinishProgress
    ModLog.LogActionEnd strAction
    
    Exit Sub

GetPatientDataFromDatabaseError:

    ModProgress.FinishProgress
    ModMessage.ShowMsgBoxError "Kan patient " & strHospNum & " niet openenen"
    ModLog.LogError Err, Err.Description
    
End Sub
