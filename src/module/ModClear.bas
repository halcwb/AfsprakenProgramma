Attribute VB_Name = "ModClear"
Option Explicit

Private Sub ClearContentsSheetRange(shtSheet As Worksheet, strRange As String)

    Dim blnLog As Boolean
    Dim strError As String
    Dim blnIsDevelop As Boolean
    Dim strPw As String
    
    On Error GoTo ClearContentError
    
    blnIsDevelop = ModSetting.IsDevelopmentMode()
    strPw = ModConst.CONST_PASSWORD
    
    shtSheet.Unprotect strPw
    shtSheet.Visible = xlSheetVisible
    
    Application.Goto Reference:=strRange
    Selection.ClearContents
    
    If strRange = ModConst.CONST_RANGE_NEOMRI Then
        Selection.Value = 50
    End If
    
    If Not blnIsDevelop Then
        shtSheet.Visible = xlSheetVeryHidden
        shtSheet.Protect strPw
    End If
    
    Exit Sub
    
ClearContentError:
    
    strError = "ClearContentSheetRange Sheet: " & shtSheet.Name & " could  not clear content of Range: " & strRange
    ModLog.LogError strError

End Sub

Public Sub ClearLab()
    
    ClearContentsSheetRange shtPedBerLab, ModConst.CONST_RANGE_PEDLAB
    ClearContentsSheetRange shtNeoBerLab, ModConst.CONST_RANGE_NEOLAB
    
End Sub

Public Sub ClearAfspraken()

    ClearContentsSheetRange shtNeoBerAfspr, ModConst.CONST_RANGE_NEOBOOL
    ClearContentsSheetRange shtNeoBerAfspr, ModConst.CONST_RANGE_NEODATA
    ClearContentsSheetRange shtNeoBerAfspr, ModConst.CONST_RANGE_NEOMRI
    
    ClearContentsSheetRange shtPedBerAfspr, ModConst.CONST_RANGE_PEDBOOL
    ClearContentsSheetRange shtPedBerAfspr, ModConst.CONST_RANGE_PEDDATA

End Sub
