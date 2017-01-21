Attribute VB_Name = "ModContitional"
Option Explicit

Public Enum FormattingSetting
    InfoSet = 1
    WarnSet = 2
    ErrSet = 3
End Enum

Private Const constPedFormSetting = "H"

Private Function GetFormattingSettingRange(ByVal intSet As FormattingSetting) As String

    GetFormattingSettingRange = constPedFormSetting & (9 + intSet)

End Function

Public Sub ClearContitionalFormatting(ByRef objSheet As Worksheet, ByVal strRange As String)

    objSheet.Range(strRange).FormatConditions.Delete

End Sub

Public Sub SetContionalFormatting(ByRef objSheet As Worksheet, ByVal strRange As String, ByVal strFormula, ByVal enmSet As FormattingSetting, ByVal blnStop)
    
    Dim objForm As FormatCondition
    Dim strSet As String
    
    strSet = GetFormattingSettingRange(enmSet)
    
    With objSheet.Range(strRange)
        .FormatConditions.Add Type:=xlExpression, Formula1:=strFormula
        Set objForm = .FormatConditions(.FormatConditions.Count)
        objForm.Interior.Color = shtGlobSettings.Range(strSet).Interior.Color
        objForm.Font.Bold = shtGlobSettings.Range(strSet).Font.Bold
        objForm.Font.Italic = shtGlobSettings.Range(strSet).Font.Italic
        objForm.Font.Color = shtGlobSettings.Range(strSet).Font.Color
        objForm.StopIfTrue = blnStop
    End With

End Sub

Private Sub SetConditionalFormattingWarnErr(ByRef objSheet As Worksheet, ByVal strRange As String, ByVal strErr As String, ByVal strWarn As String, ByVal intStart As Integer, ByVal intStop As Integer, ByVal intOffSet As Integer)

    Dim intN As Integer

    For intN = intStart To intStop
        ClearContitionalFormatting objSheet, strRange & intN
        
        SetContionalFormatting objSheet, strRange & intN, strErr & (intN - intOffSet), ErrSet, True
        SetContionalFormatting objSheet, strRange & intN, strWarn & (intN - intOffSet), WarnSet, False
    Next
    

End Sub

Private Sub SetConditionalFormattingInfo(ByRef objSheet As Worksheet, ByVal strRange As String, ByVal strInfo As String, ByVal intStart As Integer, ByVal intStop As Integer, ByVal intOffSet As Integer)

    Dim intN As Integer

    For intN = intStart To intStop
        ClearContitionalFormatting objSheet, strRange & intN
        
        SetContionalFormatting objSheet, strRange & intN, strInfo & (intN - intOffSet), InfoSet, True
    Next
    

End Sub

Public Sub SetPedMedIVConditionalFormatting()

    Dim strWarn As String
    Dim strErr As String
    Dim strInfoSterkte As String
    Dim strInfoVol As String
    
    strErr = "=PedBerMedIV!U"
    strWarn = "=PedBerMedIV!V"
    strInfoSterkte = "=PedBerMedIV!W"
    strInfoVol = "=PedBerMedIV!X"
    
    SetConditionalFormattingInfo shtPedGuiMedIV, "F", strInfoSterkte, 9, 23, 4
    SetConditionalFormattingInfo shtPedGuiMedIV, "G", strInfoVol, 9, 23, 4
    SetConditionalFormattingWarnErr shtPedGuiMedIV, "J", strErr, strWarn, 9, 23, 4
    
End Sub

Public Sub SetNeoMedIVConditionalFormatting()

    Dim strWarn As String
    Dim strErr As String
    Dim strInfoStand As String
    Dim strInfoVol As String
    
    strErr = "=NeoBerInfB!V"
    strWarn = "=NeoBerInfB!U"
    strInfoStand = "=NeoBerInfB!R"
    strInfoVol = "=NeoBerInfB!S"
    
    SetConditionalFormattingInfo shtNeoGuiInfB, "K", strInfoStand, 28, 37, 5
    SetConditionalFormattingInfo shtNeoGuiInfB, "H", strInfoVol, 28, 37, 5
    SetConditionalFormattingWarnErr shtNeoGuiInfB, "L", strErr, strWarn, 28, 37, 5
    
    strErr = "=NeoBerInfB!X"
    SetConditionalFormattingWarnErr shtNeoGuiInfB, "C", strErr, strErr, 9, 9, 6
    
End Sub

