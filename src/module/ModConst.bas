Attribute VB_Name = "ModConst"
Option Explicit

' Number of interface sheets
Public Const CONST_INTERFACESHEET_COUNT = 13
' Number of calculation sheets
Public Const CONST_CALCULATIONSHEET_COUNT = 17

' Current name of the workbook
Public Const CONST_WORKBOOKNAME = "Afspraken2015.xlsm"
' Password to protect code and sheets
Public Const CONST_PASSWORD = "hla"

'Length bedname
Public Const CONST_BEDNAME_LENGTH As Integer = 8

'Default error message
Public Const CONST_DEFAULTERROR_MSG = "Er is een fout opgetreden. Neem contact op met uw functioneel beheerder."

'Named ranges constants
Public Const CONST_AANVULLEND_BOOLEANS = "_Aanvullend_Booleans"
Public Const CONST_AANVULLEND_DATA = "_Aanvullend_Data"
Public Const CONST_AANVULLEND_MRI_VERTREKTIJD = "_Aanvullend_MRIvertrektijd"
Public Const CONST_AANVULLEND_BOOLEANS_PED = "_Aanvullend_Booleans_Ped"
Public Const CONST_AANVULLEND_DATA_PED = "_Aanvullend_Data_Ped"
Public Const CONST_AANVULLEND_MRI_VERTREKTIJD_PED = "_Aanvullend_MRIvertrektijd_Ped"
Public Const CONST_LABDATA = "Lab_Data"
Public Const CONST_LABDATA_NEO = "LabNeo_Data"

'TPN ranges
Public Const CONST_TPN_1 As Integer = 2
Public Const CONST_TPN_2 As Integer = 7
Public Const CONST_TPN_3 As Integer = 16
Public Const CONST_TPN_4 As Integer = 30
Public Const CONST_TPN_5 As Integer = 50

' Make sure that the active workbook is Afspraken2015.xlsm
' and return the path of the Afspraken2015 workbook
Public Function GetAfsprakenProgramFilePath() As String

    Workbooks(CONST_WORKBOOKNAME).Activate
    GetAfsprakenProgramFilePath = ActiveWorkbook.Path
    
End Function

' Make sure that the active workbook is Afspraken2015.xlsm
' and return the path of the Formularium workbook
Public Function GetFormulariumDatabasePath() As String
    Dim strPath As String
    Dim arrPath() As String
    Dim intCounter As Integer

    strPath = vbNullString
    Workbooks(CONST_WORKBOOKNAME).Activate
    arrPath = Split(ActiveWorkbook.Path, "\")
    
    For intCounter = 0 To (UBound(arrPath) - 2)
        strPath = strPath & arrPath(intCounter) & "\"
    Next
    
    GetFormulariumDatabasePath = strPath & "db\"

End Function

Public Sub SetApplicationCursorToDefault()

    Application.Cursor = xlDefault

End Sub
