Attribute VB_Name = "ModTests"
Option Explicit

' Prevent display of msg box
Private blnDontDisplay As Boolean

Const blnNotFound = 0

' Run the whole test suite
Public Sub RunTests()

    ' Make sure that error messages are not displayed and
    ' program is not closed
    blnDontDisplay = True
    blnDontClose = True
    
    ' Run the tests
    Test_Open
    Test_GetPatientDataPath
    Test_GetPatientDataFile
    Test_GetPatientWorkBookName
    Test_OpenBed
    Test_CountInterfaceSheets
    Test_CountCalculationSheets
    Test_GaNaar
    Test_CanReadFormulariumDb
    Test_ClearPatient
    Test_SluitBed
    Test_Sluit
    
    MsgBox "All tests ran!", vbExclamation
    
    ' Set the program to the initial state
    Openen
    
    ' Set program to close and messages to display again
    blnDontClose = False
    blnDontDisplay = False

End Sub

' --- Individual Tests --


Public Sub Test_CanPrescribeAdrenaline()

   MsgBox shtGuiMedicatieIV.Shapes.Count

End Sub

Public Sub Test_Open()

    On Error GoTo Assert:
    Openen
    
Assert:

    CheckCursor "Opening the program did not work correctly: "

End Sub


Public Sub Test_Sluit()

    On Error GoTo Assert:
    Afsluiten

Assert:
    
    CheckCursor "Closing the program did not work correctly: "

End Sub

Public Sub Test_ClearPatient()

    Dim objCursor As XlMousePointer

    On Error GoTo Assert:
    clearPat True

Assert:

    CheckCursor "Clear patient did not work correctly: "

End Sub

Public Sub Test_GetPatientDataPath()
    
    Dim strPath As String
    Dim strFile As String
    Dim strMsg As String
    
    strPath = GetPatientDataPath()
    strFile = Dir(strPath)
    strMsg = "Patient data file could not be found in: " + strPath + " in file: " + strFile
    
    AssertNotEqual InStr(strFile, "Patient"), blnNotFound, strMsg, Not blnDontDisplay

End Sub

Public Sub Test_GetPatientDataFile()

    Dim strBed As String, strResult As String, strFile As String
    strBed = "2.9"
    
    strResult = GetPatientDataFile(strBed)
    strFile = "Patient2.9.xls"
    
    AssertNotEqual InStr(strResult, strFile), blnNotFound, "Could not get correct patient data strFile from: " + strResult, Not blnDontDisplay

End Sub

Public Sub Test_GetPatientWorkBookName()

    Dim strBed As String
    strBed = "2.9"

    AssertEqual GetPatientWorkBookName(strBed), "Patient" + strBed + ".xls", "Could not get correct CONST_WORKBOOKNAME", Not blnDontDisplay

End Sub


Public Sub Test_OpenBed()

    Dim strBed As String
    
    BeOpenenBed "2.9"
    strBed = Range("Bednummer").Formula
    
    AssertEqual "2.9", strBed, "Bed 2.9 should be opened, but strBed: " + strBed + " was open", Not blnDontDisplay

End Sub

Public Sub Test_SluitBed()

    On Error GoTo Assert:
    BeSluitBed

Assert:

    CheckCursor "Close bed did not succeed"

End Sub

Public Sub Test_CanOpenCloseWorkbook()
    
    Dim strFileName As String, strName As String, intCount As Integer
    
    intCount = Workbooks.Count
    
    strName = "Formularium.xlsx"
    strFileName = ModGlobal.GetAfsprakenProgramFilePath() + "\db\" + strName
    Workbooks.Open strFileName, True, True
    
    AssertTrue Workbooks.Count = intCount + 1, "After opening the count of workbooks should be +1", Not blnDontDisplay
    
    Workbooks(strName).Close

    AssertTrue Workbooks.Count = intCount, "After closing the count of workbooks should be original count", Not blnDontDisplay

End Sub

Public Sub Test_CanReadFormulariumDb()

    Dim objForm As clsFormularium
    Dim intCount As Integer
    Dim objMed As clsMedicatieDiscontinue
    
    intCount = 2284
    Set objForm = New clsFormularium
    
    AssertEqual intCount, objForm.MedicamentCount, "Medicament count should be: " + CStr(intCount), Not blnDontDisplay
    
    Set objMed = objForm.Item(100)
    AssertTrue objMed.Generiek <> vbNullString, "Medicament should have a generic name", Not blnDontDisplay
    
End Sub

Public Sub Test_GaNaar()

    Dim strMsg As String
    strMsg = "Not the correct sheet"

    gaNaarMedicatieIV
    AssertEqual ActiveSheet.Name, "MedicatieIV", strMsg, Not blnDontDisplay
        
    gaNaarMedicatieOverig
    AssertEqual ActiveSheet.Name, "Med_disc", strMsg, Not blnDontDisplay
    
    gaNaarInfusen
    AssertEqual ActiveSheet.Name, "Infusen", strMsg, Not blnDontDisplay
    
    gaNaarIntake
    AssertEqual ActiveSheet.Name, "Intake", strMsg, Not blnDontDisplay
    
    gaNaarLab
    AssertEqual ActiveSheet.Name, "Lab", strMsg, Not blnDontDisplay
    
    gaNaarAcuteOpvang
    AssertEqual ActiveSheet.Name, "AcuteOpvang", strMsg, Not blnDontDisplay
    
    gaNaarAfspraakBlad
    AssertEqual ActiveSheet.Name, "AfsprakenCMV", strMsg, Not blnDontDisplay
    
    gaNaarMedicatie
    AssertEqual ActiveSheet.Name, "Medicatie", strMsg, Not blnDontDisplay
    
    gaNaarTPNblad
    AssertTrue InStr(ActiveSheet.Name, "TPN") > 0, strMsg, Not blnDontDisplay

End Sub

Public Sub Test_CountInterfaceSheets()

    Dim intCount As Integer
    
    intCount = ModSheets.GetUserInterfaceSheets().Count
    
    AssertEqual intCount, CONST_INTERFACESHEET_COUNT, "Wrong number of interaces sheets", Not blnDontDisplay

End Sub

Public Sub Test_CountCalculationSheets()

    Dim intCount As Integer
    
    intCount = ModSheets.GetNonInterfaceSheets().Count
    
    AssertEqual intCount, CONST_CALCULATIONSHEET_COUNT, "Wrong number of calculation sheets", Not blnDontDisplay

End Sub

' --- Helper Functions ---

Private Function ErrorToString()

    If Err.Number <> 0 Then
        ErrorToString = "Error # " & str(Err.Number) & " was generated by " _
            & Err.Source & vbNewLine & Err.Description
    End If

End Function

Private Sub CheckCursor(strMsg As String)

    Dim objCursor As XlMousePointer
    
    objCursor = Application.Cursor
    ' Last action is set the cursor back to default
    ' So, if cursor is not default, something went wrong
    If objCursor <> xlDefault Then
        AssertTrue False, "Clear patient did not work correctly: " + ErrorToString(), Not blnDontDisplay
    End If
    
    Application.Cursor = xlDefault

End Sub

Private Sub AllNamedRanges()
Dim nm As Variant
    For Each nm In ThisWorkbook.Names
        Debug.Print nm.Name
    Next nm
End Sub

Private Sub UnhideAllSheets()
    Dim ws As Worksheet
     
    For Each ws In ActiveWorkbook.Worksheets
     
    ws.visible = xlSheetVisible
     
    Next ws
End Sub


Private Sub GetFormulas()
Dim intRow As Integer
Dim intCol As Integer

    For intRow = 7 To 26
        For intCol = 12 To 12
            If Cells(intRow, intCol).Formula <> vbNullString Then
                Debug.Print Cells(intRow, intCol).Formula
            End If
        Next
    Next
End Sub

Private Sub ChangesNamedRangeScopes2Workbook()
Dim nm As Variant
Dim strName As String
Dim strRange As String
    Dim s As String

    For Each nm In ThisWorkbook.Names
        If LCase(Left(nm.Name, 8)) = "ber_lab!" Then
            strName = nm.Name
            strRange = nm.RefersTo
            
            s = Split(nm.Name, "!")(UBound(Split(nm.Name, "!")))
            ' Add to "Workbook" scope
            nm.RefersToRange.Name = s
            ' Remove from "Worksheet" scope
            Debug.Print nm.Name & " - " & nm.RefersTo
            Call nm.Delete
        End If
    Next nm
End Sub

Private Sub UnlockAllA1Cells()
Dim ws As Worksheet

    For Each ws In ActiveWorkbook.Sheets
        Debug.Print ws.Name
        ws.Unprotect CONST_PASSWORD
        ws.visible = xlSheetVisible
        ws.Range("A6").Locked = False
        ws.Protect CONST_PASSWORD
    Next
End Sub

Private Sub TestPatientVerhuizen()

    Dim strFileName As String, strBookName As String, strRange As String, strBed As String, strTekstFile, strTekstBookName
    Dim strFileNameOld As String, strBookNameOld As String, strTekstFileOld, strTekstBookNameOld
    Dim strPrompt As String, varReply As Variant
    Dim intCount As Integer
    
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
                Else
                    Set colPatienten = Nothing
                    Set oFrmPatientLijst = Nothing
                    Exit Sub
                End If
            End With
        End If
    End If

    LogActionEnd "BeSluitBed"
    Application.Cursor = xlDefault
    
End Sub

Sub testOpenPatient()
    Dim strIndex As String
    Dim objPatienten
    Dim intCount As Integer
    
    Set objPatienten = New Collection
    
    Set objPatienten = oPuPatientenCollectie
    
    With frmPatLijst
        Application.Cursor = xlWait
        .lstPatienten.Clear
        For intCount = 1 To objPatienten.Count
            .lstPatienten.AddItem objPatienten(intCount)
        Next intCount
        Application.Cursor = xlDefault
        .Show
        If .lstPatienten.ListIndex > -1 Then
            Application.Cursor = xlWait
            strIndex = VBA.Left$(.lstPatienten.Text, CONST_BEDNAME_LENGTH) 'intBednameLength=8
            Call BeOpenenBed(strIndex)
            Application.Cursor = xlDefault
        End If
        .lstPatienten.Clear
    End With
    
    Set objPatienten = Nothing
End Sub

Public Sub testBeOpenenBed(strBed As String)
On Error GoTo BeOpenenBedError

    Dim strAction As String, strParams() As Variant
    Dim intCount As Integer
    
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

Sub GetAllNamedRangesOnCurrentWorksheet()
Dim curSheet As Worksheet
Dim x As Name
    Set curSheet = ActiveSheet
    
    For Each x In ActiveWorkbook.Names
        If InStr(1, x.Name, "Werkbrief") > 0 Then
            Debug.Print x.Name
        End If
    Next
    
End Sub

Sub TestTakenMetaVision()
Dim strTaak As String
Dim strAfspraak As String
Dim strTaak2 As String
Dim strAfspraak2 As String

strTaak = "paracetamol||2 dd||120 mg (24 mg/kg/dag)||or||||||"
strAfspraak = "paracetamol||2 dd||120 mg (30 mg/kg/dag)||or||||||"

strTaak2 = Left(strTaak, InStr(1, strTaak, "(") - 1)
strAfspraak2 = Left(strAfspraak, InStr(1, strAfspraak, "(") - 1)

If strTaak2 = strAfspraak2 Then MsgBox "Hetzelfde!"
End Sub

Sub TestVerwijderen()
    clearPat True
End Sub