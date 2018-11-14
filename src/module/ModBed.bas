Attribute VB_Name = "ModBed"
Option Explicit

Private Const constVersie As String = "Var_Glob_Versie"
Private Const constBusy As String = "DB_DatabaseBusy"

Private Const constBed As String = "__1_Bed"
Private Const constHospNum As String = "__0_PatNum"

Public Sub SetPatientHospitalNumber(ByVal strHospNum As String)

    ModRange.SetRangeValue constHospNum, strHospNum

End Sub


Public Sub SetBed(ByVal strBed As String)

    ModRange.SetRangeValue constBed, strBed

End Sub

Public Sub SetFileVersie(ByVal dtmVersie As Date)

    ModRange.SetRangeValue constVersie, dtmVersie

End Sub

Public Function GetFileVersie() As Date

    GetFileVersie = ModRange.GetRangeValue(constVersie, Now())

End Function

Public Function GetDatabaseVersie() As String

    Dim strVer As String
    
    strVer = ModRange.GetRangeValue(constVersie, "")
    If Not strVer = vbNullString Then strVer = ModDate.FormatDateTimeSeconds(CDate(strVer))
    
    GetDatabaseVersie = strVer

End Function

Private Sub Test_GetDatabaseVersie()

    ModMessage.ShowMsgBoxOK GetDatabaseVersie()

End Sub


Public Sub SetDatabaseVersie(strVersie As String)

    ModRange.SetRangeValue constVersie, strVersie

End Sub

Public Function GetBed() As String

    GetBed = ModRange.GetRangeValue(constBed, vbNullString)

End Function

Public Sub OpenBed()

    ModProgress.StartProgress "Open Bed"
    OpenBedAsk True, True
    ModProgress.FinishProgress

End Sub

Public Sub OpenBed2()

    ModProgress.StartProgress "Open Bed"
    OpenBedAsk2 True, True
    ModProgress.FinishProgress

End Sub

Private Function IsValidBed(ByVal strBed As String) As Boolean

    Dim varItem As Variant
    
    For Each varItem In ModSetting.GetPedBeds()
        If CStr(varItem) = strBed Then
            IsValidBed = True
            Exit Function
        End If
    Next varItem
    
    For Each varItem In ModSetting.GetNeoBeds()
        If CStr(varItem) = strBed Then
            IsValidBed = True
            Exit Function
        End If
    Next varItem
    
    IsValidBed = False

End Function

Public Sub OpenBedAsk(ByVal blnAsk As Boolean, ByVal blnShowProgress As Boolean)
    
    On Error GoTo ErrorOpenBed

    Dim strBed As String
    Dim strTitle As String
    Dim strAction As String
    Dim strParams() As Variant
    Dim strFileName As String
    Dim strBookName As String
    Dim strRange As String
    Dim blnAll As Boolean
    Dim blnNeo As Boolean
    
    strBed = GetBed()
    blnNeo = MetaVision_IsNeonatologie()
    
    If blnAsk Then
        If blnShowProgress Then
            strTitle = FormProgress.Caption
            ModProgress.FinishProgress
        End If
    
        ModPatient.OpenPatientLijst "Selecteer een patient"
        If GetBed() = vbNullString Then ' No bed was selected
            SetBed strBed               ' Put back the old bed
            Exit Sub                    ' And exit sub
        Else
            strBed = GetBed()
            
            If blnShowProgress Then ModProgress.StartProgress strTitle
        End If
    End If
    
    If Not IsValidBed(strBed) Then Exit Sub
            
    strAction = "ModBed.OpenBed"
    strParams = Array(strBed)
    
    ModLog.LogActionStart strAction, strParams
    
    strFileName = ModSetting.GetPatientDataFile(strBed)
    strBookName = ModSetting.GetPatientDataWorkBookName(strBed)
    strRange = "A1"
    
    If ModWorkBook.CopyWorkbookRangeToSheet(strFileName, strBookName, strRange, shtGlobTemp, True) Then
        SetFileVersie FileSystem.FileDateTime(strFileName)
        SetBed strBed
        
        If shtGlobTemp.Range("A1").CurrentRegion.Rows.Count <= 1 Then
            ModPatient.ClearPatientData vbNullString, False, True               ' Patient data was empty so clean all current data
            SetBed strBed
        End If
        
        blnAll = ModRange.CopyTempSheetToNamedRanges(True)
        If blnNeo Then ModNeoInfB.CopyCurrentInfDataToVar blnShowProgress       ' Make sure that infuusbrief data is updated
        
        If Not blnAll And blnAsk Then
            If blnShowProgress Then
                strTitle = FormProgress.Caption
                ModProgress.FinishProgress
            End If
        
            ModMessage.ShowMsgBoxExclam "Niet alle data kon worden teruggezet!" & vbNewLine & "Controleer de afspraken goed"
        
            If blnShowProgress Then ModProgress.StartProgress strTitle
            
            ModApplication.SetApplicationTitle
        End If
    End If

    ModMetaVision.MetaVision_SyncLab
    ModSheet.SelectPedOrNeoStartSheet (Not blnShowProgress)
    
    ModLog.LogActionEnd strAction
    
    Exit Sub

ErrorOpenBed:

    ModMessage.ShowMsgBoxError "Kan bed " & strBed & " niet openenen"
    ModLog.LogError Err.Description
    
End Sub

Private Sub TestOpenBed()

    OpenBed

End Sub

Public Sub OpenBedAsk2(ByVal blnAsk As Boolean, ByVal blnShowProgress As Boolean)
    
    Dim strTitle As String
    Dim strAction As String
    Dim strParams() As Variant
    Dim strFileName As String
    Dim strBookName As String
    Dim strRange As String
    Dim blnAll As Boolean
    Dim blnNeo As Boolean
    Dim strHospNum As String
    
    Dim objPat As ClassPatientDetails
    
    On Error GoTo ErrorOpenBed
    
    Set objPat = New ClassPatientDetails
    ModMetaVision.MetaVision_GetPatientDetails objPat, ModMetaVision.MetaVision_GetCurrentPatientID, vbNullString
    blnNeo = MetaVision_IsNeonatologie()
    
    strHospNum = objPat.PatientId
    If blnAsk And strHospNum = vbNullString Then
        If blnShowProgress Then
            strTitle = FormProgress.Caption
            ModProgress.FinishProgress
        End If
    
        ModPatient.OpenPatientLijst2 "Selecteer een patient"
        If ModRange.GetRangeValue(constHospNum, vbNullString) = vbNullString Then  ' No patient was selected
            SetPatientHospitalNumber strHospNum                                    ' Put back the old hospital number
            Exit Sub                                                               ' And exit sub
        Else
            strHospNum = ModRange.GetRangeValue(constHospNum, vbNullString)
            
            If blnShowProgress Then ModProgress.StartProgress strTitle
        End If
    End If
    
    GetPatientDataFromDatabase strHospNum
    If Not strHospNum = vbNullString Then SetPatientHospitalNumber strHospNum
    
    Exit Sub

ErrorOpenBed:

    ModMessage.ShowMsgBoxError "Kan bed " & strHospNum & " niet openenen"
    ModLog.LogError Err.Description
    
End Sub

Private Sub Test_OpenBedAsk2()
    
    ModProgress.StartProgress "Testing select patient"
    OpenBedAsk2 True, True
    ModProgress.FinishProgress

End Sub

Public Sub GetPatientDataFromDatabase(ByVal strHospNum As String)
    
    On Error GoTo GetPatientDataFromDatabaseError

    Dim strTitle As String
    Dim strAction As String
    Dim strParams() As Variant
    Dim blnNeo As Boolean
    
    ModProgress.StartProgress "Patient data ophalen voor " & strHospNum
    
    blnNeo = MetaVision_IsNeonatologie()
    
    strAction = "ModBed.GetPatientDataFromDatabase"
    
    ModLog.LogActionStart strAction, strParams
            
    ModPatient.ClearPatientData vbNullString, False, True
    
    ModDatabase.Database_GetPatientData strHospNum
    
    If blnNeo Then ModNeoInfB.CopyCurrentInfDataToVar True       ' Make sure that infuusbrief data is updated
            
    ModApplication.SetApplicationTitle

    ModMetaVision.MetaVision_SyncLab
    ModSheet.SelectPedOrNeoStartSheet True
    
    ModProgress.FinishProgress
    ModLog.LogActionEnd strAction
    
    Exit Sub

GetPatientDataFromDatabaseError:

    ModProgress.FinishProgress
    ModMessage.ShowMsgBoxError "Kan patient " & strHospNum & " niet openenen"
    ModLog.LogError Err.Description
    
End Sub

Private Sub Test_GetPatientDataFromDatabase()

    GetPatientDataFromDatabase "0250574"

End Sub

Public Sub CloseBed(ByVal blnAsk As Boolean)

    Dim strBed As String
    Dim strNew As String
    
    Dim strPrompt As String
    Dim strAction As String
    Dim strParams() As Variant
    
    Dim varReply As VbMsgBoxResult
    
    Dim blnNeo As Boolean

    On Error GoTo CloseBedError
    
    strBed = GetBed()
    blnNeo = MetaVision_IsNeonatologie()
    
    strAction = "ModBed.CloseBed"
    strParams = Array(blnAsk, strBed)
    LogActionStart strAction, strParams
    
    If strBed = vbNullString Then
        If blnAsk Then     ' No bed selected so ask for a bed
            ModPatient.OpenPatientLijst "Selecteer een bed"
            CloseBed False ' And try again, but do not ask again
            Exit Sub
        Else               ' No bed selected do not ask, so exit
            Exit Sub
        End If
    End If
    
    If blnAsk Then
        strPrompt = "Patient " & ModPatient.GetPatientString() & " opslaan op bed: " & strBed & "?"
        varReply = ModMessage.ShowMsgBoxYesNo(strPrompt)
    Else
        varReply = vbYes
    End If
    

    If varReply = vbYes Then
        ModProgress.StartProgress "Bed Opslaan"
        
        If blnNeo Then
            ModNeoInfB.NeoInfB_SelectInfB False, False ' Make sure that the Infuusbrief Actueel is selected
            ModNeoInfB.CopyCurrentInfVarToData True   ' Make sure that neo data is updated with latest current infuusbrief
        End If
    
        If SaveBedToFile(strBed, False, True) Then
            ModProgress.FinishProgress
            ModMessage.ShowMsgBoxInfo "Patient is opgeslagen op bed: " & strBed
        Else
            ModProgress.FinishProgress
            ModMessage.ShowMsgBoxExclam "Patient werd niet opgeslagen"
        End If
    Else
        varReply = ModMessage.ShowMsgBoxYesNo("Op een ander bed opslaan?")
        
        If varReply = vbYes Then
            ModPatient.OpenPatientLijst "Selecteer een bed"
            
            strNew = GetBed()
            ModProgress.StartProgress "Verplaats Patient Naar Bed: " & strNew
            
            If Not strNew = vbNullString And SaveBedToFile(strNew, True, True) Then
                If strBed <> vbNullString And strBed <> strNew Then
                    SetBed strBed
                    OpenBedAsk False, True
                    
                    ModPatient.PatientClearAll False, True
                    SaveBedToFile strBed, True, True
                    
                    SetBed strNew
                    OpenBedAsk False, True
                    
                    ModProgress.FinishProgress
                    ModMessage.ShowMsgBoxInfo "Patient is overgeplaatst van bed: " & strBed & " naar bed: " & strNew
                Else
                    ModProgress.FinishProgress
                    ModMessage.ShowMsgBoxInfo "Patient is opgeslagen op bed: " & strBed
                End If
            
            Else
                ModProgress.FinishProgress
                
                If strNew = vbNullString Then
                    SetBed strBed
                    ModMessage.ShowMsgBoxExclam "Patient werd niet opgeslagen"
                Else
                    ModMessage.ShowMsgBoxExclam "Patient kon niet worden opgeslagen op bed: " & strNew
                End If
            End If
        End If
    End If

    LogActionEnd strAction
    
    Exit Sub
    
CloseBedError:

    ModProgress.FinishProgress
    
    ModMessage.ShowMsgBoxError "Kan patient niet opslaan op bed: " & strBed
    ModLog.LogError strAction

End Sub

Private Function SaveBedToFile(ByVal strBed As String, ByVal blnForce As Boolean, ByVal blnShowProgress As Boolean) As Boolean

    Dim strDataRange As String
    Dim strTextRange As String
    Dim dtmVersion As Date
    Dim dtmCurrent As Date
    Dim strMsg As String
    Dim varReply As VbMsgBoxResult
    
    Dim strDataFile As String
    Dim strDataName As String
    Dim strTextFile As String
    Dim strTextName As String
    
    Dim strProg As String
        
    On Error GoTo SaveBedToFileError
    
    ' Guard for invalid bed name
    If Not IsValidBed(strBed) Then GoTo SaveBedToFileError
    
    strDataFile = ModSetting.GetPatientDataFile(strBed)
    strTextFile = ModSetting.GetPatientTextFile(strBed)
    
    ' Guard for non existing files
    If Not ModFile.FileExists(strDataFile) Or Not ModFile.FileExists(strTextFile) Then GoTo SaveBedToFileError
    
    strDataName = ModSetting.GetPatientDataWorkBookName(strBed)
    strTextName = ModSetting.GetPatientTextWorkBookName(strBed)
    
    dtmVersion = FileSystem.FileDateTime(strDataFile)
    dtmCurrent = ModBed.GetFileVersie()
    
    If blnShowProgress Then ModProgress.SetJobPercentage "Bestand Opslaan", 100, 33
    
    If Not dtmVersion = dtmCurrent And Not blnForce Then
        If blnShowProgress Then
            strProg = FormProgress.Caption
            ModProgress.FinishProgress
        End If
    
        strMsg = strMsg & "De afspraken zijn inmiddels gewijzig!" & vbNewLine
        strMsg = strMsg & "Wilt u toch de afspraken opslaan?"
        varReply = ModMessage.ShowMsgBoxYesNo(strMsg)
        
        If blnShowProgress Then ModProgress.StartProgress strProg

        If varReply = vbNo Then Exit Function
    End If
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
        
    strDataRange = "A1:B" + CStr(shtPatData.Range("B1").CurrentRegion.Rows.Count)
    strTextRange = "A1:C" + CStr(shtPatText.Range("C1").CurrentRegion.Rows.Count)
    
    FileSystem.SetAttr strDataFile, Attributes:=vbNormal ' Open Patient Data File
    Application.Workbooks.Open strDataFile, True
    
    With Workbooks(strDataName) ' Save Patient Data
        .Sheets(1).Cells.Clear
        shtPatData.Range(strDataRange).Copy
        .Sheets(1).Range("A1").PasteSpecial xlPasteValues
        .Save
        .Close
    End With
    SetFileVersie FileSystem.FileDateTime(strDataFile)
    
    If blnShowProgress Then ModProgress.SetJobPercentage "Bestand Opslaan", 100, 66
    
    FileSystem.SetAttr strTextFile, Attributes:=vbNormal ' Open Patient Text File
    Application.Workbooks.Open strTextFile, True
    
    With Workbooks(strTextName) ' Save Patient Text
        .Sheets(1).Cells.Clear
        shtPatText.Range(strTextRange).Copy
        .Sheets(1).Range("A1").PasteSpecial xlPasteValues
        .Save
        .Close
    End With
            
    If blnShowProgress Then ModProgress.SetJobPercentage "Bestand Opslaan", 100, 100
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
        
    SaveBedToFile = True
        
    Exit Function
    
SaveBedToFileError:

    ModLog.LogError "Could not save bed to files with: " & Join(Array(strBed, strDataFile, strTextFile), ", ")

    Application.DisplayAlerts = True
    SaveBedToFile = False

End Function

