Attribute VB_Name = "ModBed"
Option Explicit

Private Const constBusy As String = "DB_DatabaseBusy"
Private Const constBed As String = "__1_Bed"

Public Sub Bed_SetBed(ByVal strBed As String)

    ModRange.SetRangeValue constBed, strBed

End Sub

Public Sub Bed_SetFileVersion(ByVal dtmVersion As Date)

    ModRange.SetRangeValue CONST_PRESCRIPTIONS_VERSION, dtmVersion

End Sub

Public Function Bed_GetFileVersion() As Date

    Bed_GetFileVersion = ModRange.GetRangeValue(CONST_PRESCRIPTIONS_VERSION, Now())

End Function

Public Function Bed_PrescriptionsVersionGet() As Integer

    Dim intVersion As String
    
    intVersion = IIf(ModRange.GetRangeValue(CONST_PRESCRIPTIONS_VERSION, 0) = vbNullString, 0, ModRange.GetRangeValue(CONST_PRESCRIPTIONS_VERSION, 0))
    
    Bed_PrescriptionsVersionGet = intVersion

End Function

Private Sub Test_Bed_PrescriptionsVersionGet()

    ModMessage.ShowMsgBoxOK Bed_PrescriptionsVersionGet()

End Sub

Public Sub Bed_PrescriptionsVersionSet(intVersion As Integer)

    ModRange.SetRangeValue CONST_PRESCRIPTIONS_VERSION, intVersion

End Sub

Public Function Bed_GetBedName() As String

    Bed_GetBedName = ModRange.GetRangeValue(constBed, vbNullString)

End Function

Public Sub Bed_OpenBed()

    ModProgress.StartProgress "Open Bed"
    Bed_OpenBedAndAsk True, True
    ModProgress.FinishProgress

End Sub

Private Function Util_IsValidBed(ByVal strBed As String) As Boolean

    Dim varItem As Variant
    
    For Each varItem In ModSetting.GetPedBeds()
        If CStr(varItem) = strBed Then
            Util_IsValidBed = True
            Exit Function
        End If
    Next varItem
    
    For Each varItem In ModSetting.GetNeoBeds()
        If CStr(varItem) = strBed Then
            Util_IsValidBed = True
            Exit Function
        End If
    Next varItem
    
    Util_IsValidBed = False

End Function

Public Sub Bed_OpenBedAndAsk(ByVal blnAsk As Boolean, ByVal blnShowProgress As Boolean)
    
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
    
    strBed = Bed_GetBedName()
    blnNeo = MetaVision_IsNICU()
    
    If blnAsk Then
        If blnShowProgress Then
            strTitle = FormProgress.Caption
            ModProgress.FinishProgress
        End If
    
        ModPatient.Patient_OpenFileList "Selecteer een patient"
        If Bed_GetBedName() = vbNullString Then ' No bed was selected
            Bed_SetBed strBed               ' Put back the old bed
            Exit Sub                    ' And exit sub
        Else
            strBed = Bed_GetBedName()
            
            If blnShowProgress Then ModProgress.StartProgress strTitle
        End If
    End If
    
    If Not Util_IsValidBed(strBed) Then Exit Sub
            
    strAction = "ModBed.Bed_OpenBed"
    strParams = Array(strBed)
    
    ModLog.LogActionStart strAction, strParams
    
    strFileName = ModSetting.GetPatientDataFile(strBed)
    strBookName = Setting_GetPatientDataWorkBookName(strBed)
    strRange = "A1"
    
    If ModWorkBook.CopyWorkbookRangeToSheet(strFileName, strBookName, strRange, shtGlobTemp, True) Then
        Bed_SetFileVersion FileSystem.FileDateTime(strFileName)
        Bed_SetBed strBed
        
        If shtGlobTemp.Range("A1").CurrentRegion.Rows.Count <= 1 Then
            ModPatient.Patient_ClearData vbNullString, False, True               ' Patient data was empty so clean all current data
            Bed_SetBed strBed
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
            
            ModApplication.App_SetApplicationTitle
        End If
    End If

    ModMetaVision.MetaVision_SyncLab
    ModSheet.SelectPedOrNeoStartSheet (Not blnShowProgress)
    
    ModLog.LogActionEnd strAction
    
    Exit Sub

ErrorOpenBed:

    ModMessage.ShowMsgBoxError "Kan bed " & strBed & " niet openenen"
    ModLog.LogError Err, Err.Description
    
End Sub

Private Sub Test_Bed_OpenBed()

    Bed_OpenBed

End Sub

Public Sub Bed_Import(ByVal strHospNum As String)

    Dim strFile As String
    Dim strMsg As String
    Dim strBed As String
    Dim strBookName As String
    Dim strRange As String
    Dim strTitle As String
    Dim objWbk As Workbook
    Dim objPat As ClassPatientDetails
    Dim blnAll As Boolean
    Dim blnNeo As Boolean
        
    strFile = ModFile.GetFileWithDialog(WbkAfspraken.Path & "/../Data")
    If strFile = "" Then Exit Sub
        
    strMsg = "Dit bestand importeren?" & vbNewLine & strFile
    If ModMessage.ShowMsgBoxYesNo(strMsg) = vbNo Then Exit Sub
    
    Set objWbk = Workbooks.Open(strFile, True, True)
    strBookName = objWbk.Name
    objWbk.Close
    
    strRange = "A1"
    
    ModProgress.StartProgress "Importeer patient"

    If ModWorkBook.CopyWorkbookRangeToSheet(strFile, strBookName, strRange, shtGlobTemp, True) Then
        Patient_ClearAll False, True
        
        strBed = ModMetaVision.MetaVision_GetPatientBed(vbNullString, strHospNum)
        Bed_SetBed strBed
        
        If shtGlobTemp.Range("A1").CurrentRegion.Rows.Count <= 1 Then
            ModPatient.Patient_ClearData vbNullString, False, True               ' Patient data was empty so clean all current data
            
            Patient_SetHospitalNumber strHospNum
            Bed_SetBed strBed
        End If
        
        blnAll = ModRange.CopyTempSheetToNamedRanges(True)
        
        Patient_SetHospitalNumber strHospNum
        Set objPat = New ClassPatientDetails
        ModMetaVision.MetaVision_GetPatientDetails objPat, vbNullString, strHospNum
        Patient_WritePatientDetails objPat, False
        
        blnNeo = MetaVision_IsNICU()
        If blnNeo Then ModNeoInfB.CopyCurrentInfDataToVar True       ' Make sure that infuusbrief data is updated
        
        If Not blnAll Then
            If True Then
                strTitle = FormProgress.Caption
                ModProgress.FinishProgress
            End If
        
            ModMessage.ShowMsgBoxExclam "Niet alle data kon worden teruggezet!" & vbNewLine & "Controleer de afspraken goed"
        
            If True Then ModProgress.StartProgress strTitle
            
            ModApplication.App_SetApplicationTitle
        End If
    End If

    ModMetaVision.MetaVision_SyncLab
    ModSheet.SelectPedOrNeoStartSheet False
    
    ModProgress.FinishProgress


End Sub

Public Sub Bed_CloseBed(ByVal blnAsk As Boolean)

    Dim strBed As String
    Dim strNew As String
    
    Dim strPrompt As String
    Dim strAction As String
    Dim strParams() As Variant
    
    Dim varReply As VbMsgBoxResult
    
    Dim blnNeo As Boolean

    On Error GoTo CloseBedError
    
    strBed = Bed_GetBedName()
    blnNeo = MetaVision_IsNICU()
    
    strAction = "ModBed.Bed_CloseBed"
    strParams = Array(blnAsk, strBed)
    LogActionStart strAction, strParams
    
    If strBed = vbNullString Then
        If blnAsk Then     ' No bed selected so ask for a bed
            ModPatient.Patient_OpenFileList "Selecteer een bed"
            Bed_CloseBed False ' And try again, but do not ask again
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
    
        If Util_SaveBedToFile(strBed, False, True) Then
            ModProgress.FinishProgress
            ModMessage.ShowMsgBoxInfo "Patient is opgeslagen op bed: " & strBed
        Else
            ModProgress.FinishProgress
            ModMessage.ShowMsgBoxExclam "Patient werd niet opgeslagen"
        End If
    Else
        varReply = ModMessage.ShowMsgBoxYesNo("Op een ander bed opslaan?")
        
        If varReply = vbYes Then
            ModPatient.Patient_OpenFileList "Selecteer een bed"
            
            strNew = Bed_GetBedName()
            ModProgress.StartProgress "Verplaats Patient Naar Bed: " & strNew
            
            If Not strNew = vbNullString And Util_SaveBedToFile(strNew, True, True) Then
                If strBed <> vbNullString And strBed <> strNew Then
                    Bed_SetBed strBed
                    Bed_OpenBedAndAsk False, True
                    
                    Patient_ClearAll False, True
                    Util_SaveBedToFile strBed, True, True
                    
                    Bed_SetBed strNew
                    Bed_OpenBedAndAsk False, True
                    
                    ModProgress.FinishProgress
                    ModMessage.ShowMsgBoxInfo "Patient is overgeplaatst van bed: " & strBed & " naar bed: " & strNew
                Else
                    ModProgress.FinishProgress
                    ModMessage.ShowMsgBoxInfo "Patient is opgeslagen op bed: " & strBed
                End If
            
            Else
                ModProgress.FinishProgress
                
                If strNew = vbNullString Then
                    Bed_SetBed strBed
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
    ModLog.LogError Err, strAction

End Sub

Private Function Util_SaveBedToFile(ByVal strBed As String, ByVal blnForce As Boolean, ByVal blnShowProgress As Boolean) As Boolean

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
    If Not Util_IsValidBed(strBed) Then GoTo SaveBedToFileError
    
    strDataFile = ModSetting.GetPatientDataFile(strBed)
    strTextFile = ModSetting.GetPatientTextFile(strBed)
    
    ' Guard for non existing files
    If Not ModFile.FileExists(strDataFile) Or Not ModFile.FileExists(strTextFile) Then GoTo SaveBedToFileError
    
    strDataName = Setting_GetPatientDataWorkBookName(strBed)
    strTextName = Setting_GetPatientTextWorkBookName(strBed)
    
    dtmVersion = FileSystem.FileDateTime(strDataFile)
    dtmCurrent = ModBed.Bed_GetFileVersion()
    
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
    ImprovePerf True
        
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
    Bed_SetFileVersion FileSystem.FileDateTime(strDataFile)
    
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
    ImprovePerf False
        
    Util_SaveBedToFile = True
        
    Exit Function
    
SaveBedToFileError:

    ModLog.LogError Err, "Could not save bed to files with: " & Join(Array(strBed, strDataFile, strTextFile), ", ")

    Application.DisplayAlerts = True
    Util_SaveBedToFile = False

End Function

