Attribute VB_Name = "ModBedden"
Option Explicit

Private intCount As Integer

Public Sub OpenBed(strBed As String)
    
    On Error GoTo ErrorOpenBed

    Dim strAction As String, strParams() As Variant
    
    strAction = "OpenenBed"
    strParams = Array(strBed)
    
    ModLogging.LogActionStart strAction, strParams
    
    Dim strFileName As String, strBookName As String, strRange As String
    
    strFileName = GetPatientDataFile(strBed)
    strBookName = GetPatientWorkBookName(strBed)
    strRange = "a1:b1"
    
    If ModPublic.CopyWorkBookRangeToTempSheet(strFileName, strBookName, strRange) Then
        Range("BedNummer").Value = strBed
        shtPedGuiLab.Unprotect (ModConst.CONST_PASSWORD)
        With shtGlobTemp
            On Error Resume Next
            For intCount = 2 To .Range("A1").CurrentRegion.Rows.Count
                Range(.Cells(intCount, 1).Value).Formula = .Cells(intCount, 2).Formula
            Next intCount
        End With
        shtPedGuiLab.Protect (ModConst.CONST_PASSWORD)
    End If

    ModMenuItems.SelectTPN

    ModLogging.LogActionEnd "BeOpenBed"
    
Exit Sub

ErrorOpenBed:
    MsgBox prompt:=ModConst.CONST_DEFAULTERROR_MSG, Buttons:=vbExclamation, Title:="Infornmedica 2000"
    Application.Cursor = xlDefault
    ModLogging.EnableLogging
    ModLogging.LogToFile ModConst.GetAfsprakenProgramFilePath() + ModSettings.GetLogDir(), Error, Err.Description
    ModLogging.DisableLogging
End Sub

Public Sub SluitBed()

    Dim strFileName As String
    Dim strFileNameOld As String
    
    Dim strBookName As String
    Dim strBookNameOld As String
    
    Dim strBed As String
    Dim strBedOld As String
    
    Dim strTekstFile As String
    Dim strTekstFileOld As String
    
    Dim strTekstBookName As String
    Dim strTekstBookNameOld As String
    
    Dim strRange As String
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

    strAction = "BeSluitBed"
    strParams = Array(strFileName, strBookName, strBed, strTekstFile, strTekstBookName)
    LogActionStart strAction, strParams
    
    strPrompt = "Patient " & Range("_VoorNaam").Value & ", " & Range("_AchterNaam") _
    & " opslaan op bed: " & strBed & "?"
    varReply = MsgBox(prompt:=strPrompt, Buttons:=vbYesNo, Title:="Informedica 200")

    If varReply = vbYes Then
        Application.Cursor = xlWait
        If SaveBedToFile(strFileName, strBookName, strTekstFile, strTekstBookName) Then
            MsgBox "Patient is opgeslagen", vbInformation, "Informedica"
        End If
        Application.Cursor = xlDefault
    Else
        varReply = MsgBox("Op een ander bed opslaan?", vbYesNo, "Informedica")
        If varReply = vbYes Then
            strBedOld = strBed
            Set colPatienten = GetPatients
            frmPatLijst.Caption = "Selecteer de patient die vervangen moet worden ..."
            With frmPatLijst.lstPatienten
                .Clear
                For intCount = 1 To colPatienten.Count
                    .AddItem colPatienten(intCount)
                Next intCount

                frmPatLijst.Show
                If .ListIndex > -1 Then
                    strBed = VBA.Left$(.Text, CONST_BEDNAME_LENGTH)
                    Range("Bednummer").Value = strBed
                    Set colPatienten = Nothing
                    ModBedden.SluitBed
                    
                    'Alleen oude verwijderen als oude bed niet 0 is
                    If strBedOld <> "0" Then
                        ModBedden.OpenBed strBedOld
                        ClearPatient False
                        'TODO: Opslaan zonder meldingen->Call BeSluitBed
                        strFileNameOld = GetPatientDataFile(strBedOld)
                        strTekstFileOld = Replace(strFileName, ".xls", "_AfsprakenTekst.xls")
    
                        strBookNameOld = "Patient" + strBedOld + ".xls"
                        strTekstBookNameOld = "Patient" + strBedOld + "_AfsprakenTekst.xls"
                        SaveBedToFile strFileNameOld, strBookNameOld, strTekstFileOld, strTekstBookNameOld
                        ModBedden.OpenBed strBed
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
                End If
            End With
        End If
    End If

    LogActionEnd "BeSluitBed"
    
    Set frmPatLijst = Nothing
    Application.Cursor = xlDefault

End Sub


Public Function GetPatientWorkBookName(strBed As String) As String

    GetPatientWorkBookName = "Patient" + strBed + ".xls"

End Function


Public Function GetPatientDataFile(strBed As String) As String

    GetPatientDataFile = GetPatientDataPath + GetPatientWorkBookName(strBed)

End Function




