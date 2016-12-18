Attribute VB_Name = "ModPatientData"
Option Explicit

Public Function GetPatientDataPath() As String

    GetPatientDataPath = GetRelativePath(CONST_PATIENT_DATAFOLDER)

End Function

Private Function GetRelativePath(strPath As String) As String

    GetRelativePath = ActiveWorkbook.Path + strPath

End Function

