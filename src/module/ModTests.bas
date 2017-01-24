Attribute VB_Name = "ModTests"
Option Explicit

' Prevent display of msg box
Private blnDontDisplay As Boolean

Const blnNotFound As Integer = 0

' Run the whole test suite
Public Sub RunTests()

    ' Make sure that error messages are not displayed and
    ' program is not closed
    blnDontDisplay = True
    ModApplication.SetDontClose True
    
    ' Run the tests
    Test_Open
    Test_GetPatientDataPath
    Test_GetPatientDataFile
    Test_GetPatientWorkBookName
    Test_OpenBed
    Test_CountInterfaceSheets
    Test_CountCalculationSheets
    Test_CanReadFormulariumDb
    Test_ClearPatient
    Test_SluitBed
    Test_Sluit
    
    MsgBox "All tests ran!", vbExclamation
    
    ' Set the program to the initial state
    ModApplication.InitializeAfspraken
    
    ' Set program to close and messages to display again
    ModApplication.SetDontClose False
    blnDontDisplay = False

End Sub

' --- Individual Tests --


Private Sub Test_CanPrescribeAdrenaline()

    MsgBox shtPedGuiMedIV.Shapes.Count

End Sub

Private Sub Test_Open()

    On Error GoTo Assert:
    
    InitializeAfspraken
    
    Exit Sub
    
Assert:

    CheckCursor "Opening the program did not work correctly: "

End Sub


Private Sub Test_Sluit()

    On Error GoTo Assert:
    CloseAfspraken
    
    Exit Sub

Assert:
    
    CheckCursor "Closing the program did not work correctly: "

End Sub

Private Sub Test_ClearPatient()

    On Error GoTo Assert:
    ModPatient.PatientClearAll True, True
    
    Exit Sub

Assert:

    CheckCursor "Clear patient did not work correctly: "

End Sub

Private Sub Test_GetPatientDataPath()
    
    Dim strPath As String
    Dim strFile As String
    Dim strMsg As String
    
    strPath = GetPatientDataPath()
    strFile = Dir(strPath)
    strMsg = "Patient data file could not be found in: " + strPath + " in file: " + strFile
    
    AssertNotEqual Strings.InStr(strFile, "Patient"), blnNotFound, strMsg, Not blnDontDisplay

End Sub

Private Sub Test_GetPatientDataFile()

    Dim strBed As String
    Dim strResult As String
    Dim strFile As String
    strBed = "2.9"
    
    strResult = GetPatientDataFile(strBed)
    strFile = "Patient2.9.xls"
    
    AssertNotEqual Strings.InStr(strResult, strFile), blnNotFound, "Could not get correct patient data strFile from: " + strResult, Not blnDontDisplay

End Sub

Private Sub Test_GetPatientWorkBookName()

    Dim strBed As String
    strBed = "2.9"

    AssertEqual GetPatientDataWorkBookName(strBed), "Patient" + strBed + ".xls", "Could not get correct CONST_WORKBOOKNAME", Not blnDontDisplay

End Sub

Private Sub Test_OpenBed()

    Dim strBed As String
    
    ModBed.SetBed "Unit 2.9"
    ModBed.OpenBed
    strBed = ModBed.GetBed
    
    AssertEqual "Unit 2.9", strBed, "Bed 2.9 should be opened, but strBed: " + strBed + " was open", Not blnDontDisplay

End Sub

Private Sub Test_SluitBed()

    On Error GoTo Assert:
    ModBed.CloseBed (False)
    
    Exit Sub

Assert:

    CheckCursor "Close bed did not succeed"

End Sub

Private Sub Test_CanReadFormulariumDb()

    Dim objForm As ClassFormularium
    Dim intCount As Integer
    Dim objMed As ClassMedicatieDisc
    
    intCount = 2284
    Set objForm = New ClassFormularium
    
    AssertEqual intCount, objForm.MedicamentCount, "Medicament count should be: " + CStr(intCount), Not blnDontDisplay
    
    Set objMed = objForm.Item(100)
    AssertTrue objMed.Generiek <> vbNullString, "Medicament should have a generic name", Not blnDontDisplay
    
End Sub

Private Sub Test_CountInterfaceSheets()

    Dim intCount As Integer
    
    intCount = ModSheet.GetUserInterfaceSheets().Count
    
    AssertEqual intCount, ModSheet.GetInterfaceSheetCount, "Wrong number of interaces sheets", Not blnDontDisplay

End Sub

Private Sub Test_CountCalculationSheets()

    Dim intCount As Integer
    
    intCount = ModSheet.GetNonInterfaceSheets().Count
    
    AssertEqual intCount, ModSheet.GetNonInterfaceSheetCount, "Wrong number of calculation sheets", Not blnDontDisplay

End Sub

' --- Helper Functions ---

Private Function ErrorToString() As String

    If err.Number <> 0 Then
        ErrorToString = "Error # " & Conversion.Str(err.Number) & " was generated by " _
                      & err.Source & vbNewLine & err.Description
    End If

End Function

Private Sub CheckCursor(ByVal strMsg As String)

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
     
        ws.Visible = xlSheetVisible
     
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

Private Sub UnlockAllA1Cells()
    Dim ws As Worksheet

    For Each ws In ActiveWorkbook.Sheets
        Debug.Print ws.Name
        ws.Unprotect CONST_PASSWORD
        ws.Visible = xlSheetVisible
        ws.Range("A6").Locked = False
        ws.Protect CONST_PASSWORD
    Next
End Sub

Private Sub GetAllNamedRangesOnCurrentWorksheet()
    Dim curSheet As Worksheet
    Dim X As Name
    Set curSheet = ActiveSheet
    
    For Each X In ActiveWorkbook.Names
        If Strings.InStr(1, X.Name, "Werkbrief") > 0 Then
            Debug.Print X.Name
        End If
    Next
    
End Sub

Private Sub TestTakenMetaVision()
    Dim strTaak As String
    Dim strAfspraak As String
    Dim strTaak2 As String
    Dim strAfspraak2 As String

    strTaak = "paracetamol||2 dd||120 mg (24 mg/kg/dag)||or||||||"
    strAfspraak = "paracetamol||2 dd||120 mg (30 mg/kg/dag)||or||||||"

    strTaak2 = Strings.Left(strTaak, Strings.InStr(1, strTaak, "(") - 1)
    strAfspraak2 = Strings.Left(strAfspraak, Strings.InStr(1, strAfspraak, "(") - 1)

    If strTaak2 = strAfspraak2 Then MsgBox "Hetzelfde!"
End Sub

Private Sub TestVerwijderen()
    ModPatient.PatientClearAll True, True
End Sub

Private Sub TestWorkBookName()

    MsgBox ActiveWorkbook.Name

End Sub

Private Sub UseByVal(ByVal strValue As String)

    strValue = "Test"

End Sub

Private Sub TestByValVsByRef()

    Dim strValue As String
    
    strValue = "Hello World"
    UseByVal strValue
    MsgBox strValue

End Sub
