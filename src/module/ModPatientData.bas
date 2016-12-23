Attribute VB_Name = "ModPatientData"
Option Explicit

Public Function GetPatientDataPath() As String

    Dim strDir As String
    
    strDir = ModSettings.GetDataDir
    GetPatientDataPath = GetRelativePath(strDir)

End Function

Private Function GetRelativePath(strPath As String) As String

    GetRelativePath = ActiveWorkbook.Path + strPath

End Function

