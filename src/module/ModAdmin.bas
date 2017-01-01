Attribute VB_Name = "ModAdmin"
Option Explicit

Private Sub SetUpDataDir(arrBeds() As Variant)
    
    Dim strPath As String
    
    strPath = ModSetting.GetPatientDataPath()
    
    Application.DisplayAlerts = False

    ModFile.DeleteAllFilesInDir strPath
    CreateDataWorkBooks arrBeds, strPath
    
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

