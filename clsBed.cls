VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_BedNummer As Variant
Private BedMap As String
Private BedOpen As Boolean

Public Sub BedOpenen(BedNummer)

On Error GoTo BedOpenError

m_BedNummer = BedNummer
BedMap = "K:\ICAP\Bed" & BedNummer & "\Patient.xls"
FileSystem.SetAttr BedMap, Attributes:=vbNormal
Application.Workbooks.Open BedMap
BedOpen = True

Exit Sub

BedOpenError:
MsgBox prompt:="Kan bed nu niet openen, probeer dadelijk nog een keer", _
 Buttons:=vbExclamation, Title:="Infornmedica 2000"
 Application.Cursor = xlDefault
 End
    ModLogging.EnableLogging
    ModLogging.LogToFile ModGlobal.GetAfsprakenProgramFilePath() + ModGlobal.CONST_LOGPATH, Info, Err.Description
    ModLogging.DisableLogging
End Sub

Public Sub CopyPatient()

If BedOpen Then
    Workbooks("Patient").Worksheets("Patienten").Range("dbPatient").Copy
    Workbooks("Afspraken").Worksheets("Patienten").Range("dbPatient").PasteSpecial
    Workbooks("Patient").Close savechanges:=True
    BedOpen = False
End If

End Sub

Public Sub BedClose(BedNummer)

On Error GoTo BedCloseError

m_BedNummer = BedNummer
BedMap = "K:\ICAP\Bed" & BedNummer & "\Patient.xls"
FileSystem.SetAttr BedMap, Attributes:=vbNormal
Application.Workbooks.Open BedMap
BedOpen = True
If BedOpen Then
    Workbooks("Afspraken").Worksheets("Patienten").Range("dbPatient").Copy
    Workbooks("Patient").Worksheets("Patienten").Range("dbPatient").PasteSpecial
    Workbooks("Patient").Close savechanges:=True
End If

Exit Sub

BedCloseError:
MsgBox prompt:="Kan bed nu niet opslaan, probeer dadelijk nog een keer", _
 Buttons:=vbExclamation, Title:="Infornmedica 2000"
 Application.Cursor = xlDefault
 End
End Sub
