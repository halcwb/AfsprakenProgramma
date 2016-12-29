Attribute VB_Name = "ModBed"
Option Explicit

Public Sub OpenBed(strBed As String)
    
    On Error GoTo ErrorOpenBed

    Dim strAction As String
    Dim strParams() As Variant
    Dim strFileName As String
    Dim strBookName As String
    Dim strRange As String
        
    strAction = "ModBed.OpenBed"
    strParams = Array(strBed)
    
    ModLog.LogActionStart strAction, strParams
    
    strFileName = ModPatient.GetPatientDataFile(strBed)
    strBookName = ModPatient.GetPatientWorkBookName(strBed)
    strRange = "a1:b1"
    
    If ModWorkBook.CopyWorkbookRangeToSheet(strFileName, strBookName, strRange, shtGlobTemp) Then
        Range(ModConst.CONST_RANGE_VERSIE).Value = FileSystem.FileDateTime(strFileName)
        Range(ModConst.CONST_RANGE_BED).Value = strBed
        
        ModRange.CopyTempSheetRangeToRange
    End If

    ModMenuItems.SelectTPN

    ModLog.LogActionEnd strAction
    
    Exit Sub

ErrorOpenBed:

    ModMessage.ShowMsgBoxError ModConst.CONST_DEFAULTERROR_MSG
    Application.Cursor = xlDefault
    
    ModLog.LogError Err.Description
    
End Sub

Public Sub SluitBed()

    Dim strFileName As String
    Dim strBookName As String
    Dim strBed As String
    Dim strTekstFile As String
    Dim strTekstBookName As String
    
    Dim strPrompt As String
    Dim strAction As String
    Dim strParams() As Variant
    
    Dim varReply As VbMsgBoxResult
    Dim colPatienten As Collection
    Dim frmPatLijst As New FormPatLijst
    
    strBed = Range("Bednummer").Formula
    strFileName = GetPatientDataFile(strBed)
    strTekstFile = Replace(strFileName, ".xls", "_AfsprakenTekst.xls")

    strBookName = "Patient" + strBed + ".xls"
    strTekstBookName = "Patient" + strBed + "_AfsprakenTekst.xls"

    strAction = "ModBed.SluitBed"
    strParams = Array(strFileName, strBookName, strBed, strTekstFile, strTekstBookName)
    LogActionStart strAction, strParams
    
    strPrompt = "Patient " & Range("_VoorNaam").Value & ", " & Range("_AchterNaam") & " opslaan op bed: " & strBed & "?"
    varReply = ModMessage.ShowMsgBoxYesNo(strPrompt)

    If varReply = vbYes Then
        Application.Cursor = xlWait
        If SaveBedToFile(strFileName, strBookName, strTekstFile, strTekstBookName) Then
            ModMessage.ShowMsgBoxInfo "Patient is opgeslagen"
        End If
        Application.Cursor = xlDefault
    Else
        varReply = ModMessage.ShowMsgBoxYesNo("Op een ander bed opslaan?")
        If varReply = vbYes Then
            strBed = ModPatient.OpenPatientLijst("Selecteer een bed")
            Range("Bednummer").Value = strBed
            ModBed.SluitBed
            'Alleen oude verwijderen als oude bed niet 0 is
            If strBed <> "0" Then
                ModBed.OpenBed strBed
                ModPatient.ClearPatient False
                
                strFileName = ModPatient.GetPatientDataFile(strBed)
                strTekstFile = Strings.Replace(strFileName, ".xls", "_AfsprakenTekst.xls")
                strBookName = "Patient" + strBed + ".xls"
                strTekstBookName = "Patient" + strBed + "_AfsprakenTekst.xls"
                
                SaveBedToFile strFileName, strBookName, strTekstFile, strTekstBookName
                OpenBed strBed
            End If
        End If
    End If

    LogActionEnd strAction
    
    Set frmPatLijst = Nothing
    Application.Cursor = xlDefault

End Sub

Public Function SaveBedToFile(strFileName As String, strWorkbook As String, strTekstFile, strTekstBookName) As Boolean

    Dim strSelectie As String
    Dim strAfsprakenTekst As String
    Dim dtmVersion As Date
    Dim strMsg As String
    Dim intAnswer As Integer
    
    If Not Range("BedNummer").Value = 0 Then
        dtmVersion = FileSystem.FileDateTime(strFileName)
        If Not dtmVersion = Range(ModConst.CONST_RANGE_VERSIE).Value Then
            strMsg = strMsg & "De afspraken zijn inmiddels gewijzig!" & vbNewLine
            strMsg = strMsg & "Wilt u toch de afspraken opslaan?"
            intAnswer = ModMessage.ShowMsgBoxYesNo(strMsg)
            
            If intAnswer = vbNo Then Exit Function
        End If
    End If
    
    Application.DisplayAlerts = False
    
    If ModPatient.CopyPatientData() Then
        strSelectie = "a1:B" + CStr(shtPatData.Range("b1").CurrentRegion.Rows.Count)
        strAfsprakenTekst = "a1:c" + CStr(shtPatDataText.Range("c1").CurrentRegion.Rows.Count)
        FileSystem.SetAttr strFileName, Attributes:=vbNormal
        Application.Workbooks.Open strFileName, True
        With Workbooks(strWorkbook)
            .Sheets("Patienten").Cells.Clear
            shtPatData.Range(strSelectie).Copy
            .Sheets("Patienten").Range("A1").PasteSpecial xlPasteValues
            .Save
            .Close
        End With
        Range(ModConst.CONST_RANGE_VERSIE).Value = FileSystem.FileDateTime(strFileName)
        
        FileSystem.SetAttr strTekstFile, Attributes:=vbNormal
        Application.Workbooks.Open strTekstFile, True
        With Workbooks(strTekstBookName)
            .Sheets("AfsprakenTekst").Cells.Clear
            shtPatDataText.Range(strAfsprakenTekst).Copy
            .Sheets("AfsprakenTekst").Range("A1").PasteSpecial xlPasteValues
            .Save
            .Close
        End With
            
    End If
    
    Application.DisplayAlerts = True
        
    SaveBedToFile = True
        
    Exit Function
    
ErrFileSave:
    ModMessage.ShowMsgBoxExclam "Kan " & strFileName & " nu niet opslaan, probeer dadelijk nog een keer"
    Application.DisplayAlerts = True

End Function
