Attribute VB_Name = "ModAdmin"
Option Explicit

' ToDo add methods to setup data files and refresh patient data admin jobs

Private Sub SetUpDataDir(ByRef arrBeds() As Variant)
    
    Dim strPath As String
    
    strPath = ModSetting.GetPatientDataPath()
    
    Application.DisplayAlerts = False
    ModProgress.StartProgress "Opzetten Data Files"

    ModFile.DeleteAllFilesInDir strPath
    ModWorkBook.CreateDataWorkBooks arrBeds, strPath, True
    
    ModProgress.FinishProgress
    Application.DisplayAlerts = True

End Sub

Private Sub SetUpPedDataDir()
    
    Dim arrBeds() As Variant
    arrBeds = ModSetting.GetPedBeds()
    
    SetUpDataDir arrBeds
    
End Sub

Private Sub SetUpNeoDataDir()
    
    Dim arrBeds() As Variant
    arrBeds = ModSetting.GetNeoBeds()
    
    SetUpDataDir arrBeds
    
End Sub

