Attribute VB_Name = "ModBedden"
Option Explicit

Private intCount As Integer

Public Sub BeOpenenBed(strBed As String)
On Error GoTo BeOpenenBedError

    Dim strAction As String, strParams() As Variant
    strAction = "BeOpenenBed"
    strParams = Array(strBed)
    LogActionStart strAction, strParams
    
    Dim strFileName As String, strBookName As String, strRange As String
    
    strFileName = GetPatientDataFile(strBed)
    strBookName = GetPatientWorkBookName(strBed)
    strRange = "a1:b1"
    
    If PuRangeCopy(strFileName, strBookName, strRange) Then
        Range("BedNummer").Value = strBed
        shtGuiLab.Unprotect (CONST_PASSWORD)
        With shtBerTemp
            On Error Resume Next
            For intCount = 2 To .Range("A1").CurrentRegion.Rows.Count
                Range(.Cells(intCount, 1).Value).Formula = .Cells(intCount, 2).Formula
            Next intCount
        End With
        shtGuiLab.Protect (CONST_PASSWORD)
    End If

    SelectTPN

    LogActionEnd "BeOpenBed"
    
Exit Sub

BeOpenenBedError:
MsgBox prompt:=CONST_DEFAULTERROR_MSG, _
 Buttons:=vbExclamation, Title:="Infornmedica 2000"
 Application.Cursor = xlDefault
    ModLogging.EnableLogging
    ModLogging.LogToFile ModGlobal.GetAfsprakenProgramFilePath() + ModGlobal.CONST_LOGPATH, Error, Err.Description
    ModLogging.DisableLogging
End Sub

Public Sub BeSluitBed()

'    Dim strFileName As String, strBookName As String, strRange As String, strBed As String, strTekstFile, strTekstBookName
'    Dim strPrompt As String, varReply As Variant
'
'    Dim colPatienten As Collection
'    Dim oFrmPatientLijst As frmPatLijst
'    Dim strBedOld As String
'
'    strBed = Range("Bednummer").Formula
'    strFileName = GetPatientDataFile(strBed)
'    strTekstFile = Replace(strFileName, ".xls", "_AfsprakenTekst.xls")
'
'    strBookName = "Patient" + strBed + ".xls"
'    strTekstBookName = "Patient" + strBed + "_AfsprakenTekst.xls"
'
'    Dim strAction As String, strParams() As Variant
'    strAction = "BeSluitBed"
'    strParams = Array(strFileName, strBookName, strBed, strTekstFile, strTekstBookName)
'    LogActionStart strAction, strParams
'
'    strPrompt = "Patient " & Range("_VoorNaam").Value & ", " & Range("_AchterNaam") _
'    & " opslaan op bed: " & strBed & "?"
'    varReply = MsgBox(prompt:=strPrompt, Buttons:=vbYesNo, Title:="Informedica 200")
'
'    If varReply = vbYes Then
'        Application.Cursor = xlWait
'        If bPuBedOpslaan(strFileName, strBookName, strTekstFile, strTekstBookName) Then
'            MsgBox "Patient is opgeslagen", vbInformation, "Informedica"
'        End If
'        Application.Cursor = xlDefault
'    Else
'        varReply = MsgBox("Op een ander bed opslaan?", vbYesNo, "Informedica")
'        If varReply = vbYes Then
'            strBedOld = strBed
'            Set colPatienten = oPuPatientenCollectie
'            Set oFrmPatientLijst = New frmPatLijst
'            oFrmPatientLijst.Caption = "Selecteer de patient die vervangen moet worden ..."
'            With oFrmPatientLijst.lstPatienten
'                .Clear
'                For intCount = 1 To colPatienten.Count
'                    .AddItem colPatienten(intCount)
'                Next intCount
'
'                oFrmPatientLijst.Show
'                If .ListIndex > -1 Then
'                    strBed = VBA.Left$(.Text, CONST_BEDNAME_LENGTH)
'                    Range("Bednummer").Value = strBed
'                    Set colPatienten = Nothing
'                    Set oFrmPatientLijst = Nothing
'                    Call BeSluitBed
'                    'TODO:Patient verwijderen van oude bed
'                    'Oude strBed bewaren, oude gegevens bestanden leeg maken ==> NIEUWE SUB
'                Else
'                    Set colPatienten = Nothing
'                    Set oFrmPatientLijst = Nothing
'                    Exit Sub
'                End If
'            End With
'        End If
'    End If
'
'    LogActionEnd "BeSluitBed"
    
'*********************************************************************************
'Code voor verhuizen patient (opslaan op ander bed en verwijderen van huidige bed)
    Dim strFileName As String, strBookName As String, strRange As String, strBed As String, strTekstFile, strTekstBookName
    Dim strFileNameOld As String, strBookNameOld As String, strTekstFileOld, strTekstBookNameOld
    Dim strPrompt As String, varReply As Variant

    Dim colPatienten As Collection
    Dim oFrmPatientLijst As frmPatLijst
    Dim strBedOld As String

    strBed = Range("Bednummer").Formula
    strFileName = GetPatientDataFile(strBed)
    strTekstFile = Replace(strFileName, ".xls", "_AfsprakenTekst.xls")

    strBookName = "Patient" + strBed + ".xls"
    strTekstBookName = "Patient" + strBed + "_AfsprakenTekst.xls"

    Dim strAction As String, strParams() As Variant
    strAction = "BeSluitBed"
    strParams = Array(strFileName, strBookName, strBed, strTekstFile, strTekstBookName)
    LogActionStart strAction, strParams

    strPrompt = "Patient " & Range("_VoorNaam").Value & ", " & Range("_AchterNaam") _
    & " opslaan op bed: " & strBed & "?"
    varReply = MsgBox(prompt:=strPrompt, Buttons:=vbYesNo, Title:="Informedica 200")

    If varReply = vbYes Then
        Application.Cursor = xlWait
        If bPuBedOpslaan(strFileName, strBookName, strTekstFile, strTekstBookName) Then
            MsgBox "Patient is opgeslagen", vbInformation, "Informedica"
        End If
        Application.Cursor = xlDefault
    Else
        varReply = MsgBox("Op een ander bed opslaan?", vbYesNo, "Informedica")
        If varReply = vbYes Then
            strBedOld = strBed
            Set colPatienten = oPuPatientenCollectie
            Set oFrmPatientLijst = New frmPatLijst
            oFrmPatientLijst.Caption = "Selecteer de patient die vervangen moet worden ..."
            With oFrmPatientLijst.lstPatienten
                .Clear
                For intCount = 1 To colPatienten.Count
                    .AddItem colPatienten(intCount)
                Next intCount

                oFrmPatientLijst.Show
                If .ListIndex > -1 Then
                    strBed = VBA.Left$(.Text, CONST_BEDNAME_LENGTH)
                    Range("Bednummer").Value = strBed
                    Set colPatienten = Nothing
                    Set oFrmPatientLijst = Nothing
                    Call BeSluitBed
                    
                    'Alleen oude verwijderen als oude bed niet 0 is
                    If strBedOld <> "0" Then
                        Call BeOpenenBed(strBedOld)
                        Call clearPat(False)
                        'TODO: Opslaan zonder meldingen->Call BeSluitBed
                        strFileNameOld = GetPatientDataFile(strBedOld)
                        strTekstFileOld = Replace(strFileName, ".xls", "_AfsprakenTekst.xls")
    
                        strBookNameOld = "Patient" + strBedOld + ".xls"
                        strTekstBookNameOld = "Patient" + strBedOld + "_AfsprakenTekst.xls"
                        bPuBedOpslaan strFileNameOld, strBookNameOld, strTekstFileOld, strTekstBookNameOld
                        Call BeOpenenBed(strBed)
                        'TODO:Patient verwijderen van oude bed
                        'Oude strBed bewaren, oude gegevens bestanden leeg maken ==> NIEUWE SUB
                        'Open old bed: Open Patient-file
                        'Clear old bed: Copy Blanc patient to file
                        'Save old bed: Save Patient-file
                        'Open old bed: Open PatientAfspraken-file
                        'Clear old bed: Copy Blanc patientAfspraken to file
                        'Save old bed: Save PatientAfspraken-file
                        'Open new bed
                    End If
                Else
                    Set colPatienten = Nothing
                    Set oFrmPatientLijst = Nothing
                    Exit Sub
                End If
            End With
        End If
    End If

    LogActionEnd "BeSluitBed"
'*********************************************************************************
    
    Application.Cursor = xlDefault

End Sub


Public Function GetPatientWorkBookName(strBed As String) As String

    GetPatientWorkBookName = "Patient" + strBed + ".xls"

End Function


Public Function GetPatientDataFile(strBed As String) As String

    GetPatientDataFile = GetPatientDataPath + GetPatientWorkBookName(strBed)

End Function




