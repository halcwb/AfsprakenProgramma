Attribute VB_Name = "ModBed"
Option Explicit

Public Sub SetBed(strBed As String)

    ModRange.SetRangeValue ModConst.CONST_RANGE_BED, strBed

End Sub

Public Function GetBed() As String

    GetBed = ModRange.GetRangeValue(ModConst.CONST_RANGE_BED, vbNullString)

End Function

Public Sub OpenBed()
    
    On Error GoTo ErrorOpenBed

    Dim strBed As String
    Dim strAction As String
    Dim strParams() As Variant
    Dim strFileName As String
    Dim strBookName As String
    Dim strRange As String
    Dim blnAll As Boolean
    
    ModPatient.OpenPatientLijst "Selecteer een patient"
    strBed = GetBed()
    
    strAction = "ModBed.OpenBed"
    strParams = Array(strBed)
    
    ModLog.LogActionStart strAction, strParams
    
    strFileName = ModSetting.GetPatientDataFile(strBed)
    strBookName = ModSetting.GetPatientDataWorkBookName(strBed)
    strRange = "A1"
    
    If ModWorkBook.CopyWorkbookRangeToSheet(strFileName, strBookName, strRange, shtGlobTemp) Then
        ModRange.SetRangeValue ModConst.CONST_RANGE_VERSIE, FileSystem.FileDateTime(strFileName)
        ModRange.SetRangeValue ModConst.CONST_RANGE_BED, strBed
        
        blnAll = ModRange.CopyTempSheetToNamedRanges()
        If Not blnAll Then ModMessage.ShowMsgBoxExclam "Niet alle data kon worden teruggezet!" & vbNewLine & "Controleer de afspraken goed"
    End If

    ModMenuItems.SelectTPN

    ModLog.LogActionEnd strAction
    
    Exit Sub

ErrorOpenBed:

    ModMessage.ShowMsgBoxError ModConst.CONST_DEFAULTERROR_MSG
    ModLog.LogError Err.Description
    
End Sub

Private Sub TestOpenBed()

    OpenBed

End Sub

Public Sub SluitBed()

    Dim strBed As String
    Dim strNew As String
    
    Dim strPrompt As String
    Dim strAction As String
    Dim strParams() As Variant
    
    Dim varReply As VbMsgBoxResult
    Dim colPatienten As Collection
    Dim frmPatLijst As New FormPatLijst
    
    On Error GoTo CloseBedError
    
    strBed = GetBed()
    
    If strBed = vbNullString Then Exit Sub
    
    strAction = "ModBed.SluitBed"
    strParams = Array()
    LogActionStart strAction, strParams
    
    strPrompt = "Patient " & ModPatient.GetPatientString() & " Opslaan op Bed: " & strBed & "?"
    varReply = ModMessage.ShowMsgBoxYesNo(strPrompt)

    If varReply = vbYes Then
        Application.Cursor = xlWait
        
        If SaveBedToFile(False) Then
            ModMessage.ShowMsgBoxInfo "Patient is opgeslagen"
        End If
        
        Application.Cursor = xlDefault
    Else
        varReply = ModMessage.ShowMsgBoxYesNo("Op een ander bed opslaan?")
        
        If varReply = vbYes Then
            ModPatient.OpenPatientLijst "Selecteer een bed"
            
            If SaveBedToFile(False) Then
                ModMessage.ShowMsgBoxInfo "Patient is opgeslagen"
                'Alleen oude verwijderen als oude bed niet "" is
                If strBed <> vbNullString Then
                    strNew = GetBed()
                    SetBed strBed
                    OpenBed
                    ModPatient.ClearPatient False
                    
                    SaveBedToFile True
                    SetBed strNew
                    OpenBed
                End If
            End If
        End If
    End If

    LogActionEnd strAction
    
    Set frmPatLijst = Nothing
    Application.Cursor = xlDefault
    
    Exit Sub
    
CloseBedError:

    Application.Cursor = xlDefault
    
    ModMessage.ShowMsgBoxError ModConst.CONST_DEFAULTERROR_MSG & vbNewLine & "Kan patient niet opslaan op bed: " & strBed
    ModLog.LogError strAction

End Sub

Public Function SaveBedToFile(blnForce As Boolean) As Boolean

    Dim strDataRange As String
    Dim strTextRange As String
    Dim dtmVersion As Date
    Dim dtmCurrent As Date
    Dim strMsg As String
    Dim varReply As VbMsgBoxResult
    
    Dim strBed As String
    Dim strDataFile As String
    Dim strDataName As String
    Dim strTextFile As String
    Dim strTextName As String
        
    On Error GoTo SaveBedToFileError
    
    strBed = GetBed()
    If strBed = vbNullString Then Exit Function
    
    strDataFile = ModSetting.GetPatientDataFile(strBed)
    strTextFile = ModSetting.GetPatientTextFile(strBed)

    strDataName = ModSetting.GetPatientDataWorkBookName(strBed)
    strTextName = ModSetting.GetPatientTextWorkBookName(strBed)
    
    dtmVersion = FileSystem.FileDateTime(strDataFile)
    dtmCurrent = ModRange.GetRangeValue(ModConst.CONST_RANGE_VERSIE, Now())
    
    If Not dtmVersion = dtmCurrent And Not blnForce Then
        strMsg = strMsg & "De afspraken zijn inmiddels gewijzig!" & vbNewLine
        strMsg = strMsg & "Wilt u toch de afspraken opslaan?"
        varReply = ModMessage.ShowMsgBoxYesNo(strMsg)
        
        If varReply = vbNo Then Exit Function
    End If
    
    Application.DisplayAlerts = False
    
    ModPatient.CopyPatientData
    
    strDataRange = "A1:B" + CStr(shtPatData.Range("B1").CurrentRegion.Rows.Count)
    strTextRange = "A1:C" + CStr(shtPatDataText.Range("C1").CurrentRegion.Rows.Count)
    
    FileSystem.SetAttr strDataFile, Attributes:=vbNormal
    Application.Workbooks.Open strDataFile, True
    
    With Workbooks(strDataName)
        .Sheets(1).Cells.Clear
        shtPatData.Range(strDataRange).Copy
        .Sheets(1).Range("A1").PasteSpecial xlPasteValues
        .Save
        .Close
    End With
    ModRange.SetRangeValue ModConst.CONST_RANGE_VERSIE, FileSystem.FileDateTime(strDataFile)
    
    FileSystem.SetAttr strTextFile, Attributes:=vbNormal
    Application.Workbooks.Open strTextFile, True
    With Workbooks(strTextName)
        .Sheets(1).Cells.Clear
        shtPatDataText.Range(strTextRange).Copy
        .Sheets(1).Range("A1").PasteSpecial xlPasteValues
        .Save
        .Close
    End With
            
    Application.DisplayAlerts = True
        
    SaveBedToFile = True
        
    Exit Function
    
SaveBedToFileError:

    ModMessage.ShowMsgBoxExclam "Kan " & strDataFile & " nu niet opslaan, probeer dadelijk nog een keer"
    Application.DisplayAlerts = True
    SaveBedToFile = False

End Function
