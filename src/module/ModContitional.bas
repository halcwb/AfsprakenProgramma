Attribute VB_Name = "ModContitional"
Option Explicit

Public Enum FormattingSetting
    InfoSet = 1
    WarnSet = 2
    ErrSet = 3
End Enum

Private Const constPedFormSetting As String = "H"

Private Function GetFormattingSettingRange(ByVal intSet As FormattingSetting) As String

    GetFormattingSettingRange = constPedFormSetting & (9 + intSet)

End Function

Public Sub ClearContitionalFormatting(objSheet As Worksheet, ByVal strRange As String)

    objSheet.Range(strRange).FormatConditions.Delete

End Sub

Public Sub SetContionalFormatting(objSheet As Worksheet, ByVal strRange As String, ByVal strFormula As String, ByVal enmSet As FormattingSetting, ByVal blnStop As Boolean)
    
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

Private Sub SetConditionalFormattingWarnErr(objSheet As Worksheet, ByVal strRange As String, ByVal strErr As String, ByVal strWarn As String, ByVal intStart As Integer, ByVal intStop As Integer, ByVal intOffSet As Integer)

    Dim intN As Integer

    For intN = intStart To intStop
        ClearContitionalFormatting objSheet, strRange & intN
        
        SetContionalFormatting objSheet, strRange & intN, strErr & (intN - intOffSet), ErrSet, True
        SetContionalFormatting objSheet, strRange & intN, strWarn & (intN - intOffSet), WarnSet, False
    Next
    

End Sub

Private Sub SetConditionalFormattingWarn(objSheet As Worksheet, ByVal strRange As String, ByVal strWarn As String, ByVal intStart As Integer, ByVal intStop As Integer, ByVal intOffSet As Integer)

    Dim intN As Integer

    For intN = intStart To intStop
        ClearContitionalFormatting objSheet, strRange & intN
        
        SetContionalFormatting objSheet, strRange & intN, strWarn & (intN - intOffSet), WarnSet, True
    Next
    

End Sub

Private Sub SetConditionalFormattingErr(objSheet As Worksheet, ByVal strRange As String, ByVal strErr As String, ByVal intStart As Integer, ByVal intStop As Integer, ByVal intOffSet As Integer)

    Dim intN As Integer

    For intN = intStart To intStop
        ClearContitionalFormatting objSheet, strRange & intN
        
        SetContionalFormatting objSheet, strRange & intN, strErr & (intN - intOffSet), ErrSet, True
    Next
    

End Sub
Private Sub SetConditionalFormattingInfoErr(objSheet As Worksheet, ByVal strRange As String, ByVal strErr As String, ByVal strInfo As String, ByVal intStart As Integer, ByVal intStop As Integer, ByVal intOffSet As Integer)

    Dim intN As Integer

    For intN = intStart To intStop
        ClearContitionalFormatting objSheet, strRange & intN
        
        SetContionalFormatting objSheet, strRange & intN, strErr & (intN - intOffSet), ErrSet, True
        SetContionalFormatting objSheet, strRange & intN, strInfo & (intN - intOffSet), InfoSet, False
    Next
    

End Sub

Private Sub SetConditionalFormattingInfo(objSheet As Worksheet, ByVal strRange As String, ByVal strInfo As String, ByVal intStart As Integer, ByVal intStop As Integer, ByVal intOffSet As Integer)

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
    Dim strConcErr As String
    
    strErr = "=PedBerMedIV!V"
    strWarn = "=PedBerMedIV!W"
    strInfoSterkte = "=PedBerMedIV!X"
    strInfoVol = "=PedBerMedIV!Y"
    strConcErr = "=PedBerMedIV!AA"
    
    SetConditionalFormattingInfoErr shtPedGuiMedIV, "F", strConcErr, strInfoSterkte, 9, 23, 4
    SetConditionalFormattingInfo shtPedGuiMedIV, "G", strInfoVol, 9, 23, 4
    SetConditionalFormattingWarnErr shtPedGuiMedIV, "J", strErr, strWarn, 9, 23, 4
    
End Sub

Public Sub SetNeoMedIVConditionalFormatting()

    Dim strWarn As String
    Dim strErr As String
    Dim strInfoStand As String
    Dim strInfoVol As String
    Dim strErrConc As String
    
    strErr = "=NeoBerInfB!V"
    strWarn = "=NeoBerInfB!U"
    strInfoStand = "=NeoBerInfB!R"
    strInfoVol = "=NeoBerInfB!S"
    strErrConc = "=NeoBerInfB!W"
    
    SetConditionalFormattingInfo shtNeoGuiInfB, "K", strInfoStand, 28, 37, 5
    SetConditionalFormattingInfo shtNeoGuiInfB, "H", strInfoVol, 28, 37, 5
    SetConditionalFormattingWarnErr shtNeoGuiInfB, "L", strErr, strWarn, 28, 37, 5
    
    SetConditionalFormattingErr shtNeoGuiInfB, "G", strErrConc, 28, 37, 5
    
End Sub

Public Sub SetMedDiscConditionalFormatting()
    
    Dim strFreqWarn As String
    Dim strDoseWarn As String
    Dim strDoseErr As String
    
    Dim strConcErr As String
    Dim strOplErr As String
    Dim strTimeErr As String
    
    Dim strMOWarn As String
    
    strFreqWarn = "=GlobBerMedDisc!BF"
    strDoseWarn = "=GlobBerMedDisc!BL"
    strDoseErr = "=GlobBerMedDisc!BM"
    
    strConcErr = "=GlobBerMedDisc!BA"
    strOplErr = "=GlobBerMedDisc!BB"
    strTimeErr = "=GlobBerMedDisc!BC"
    
    strMOWarn = "=GlobBerMedDisc!BN"
    
    SetConditionalFormattingWarn shtGlobGuiMedDisc, "J", strFreqWarn, 9, 38, 7
    SetConditionalFormattingWarnErr shtGlobGuiMedDisc, "N", strDoseErr, strDoseWarn, 9, 38, 7
    
    SetConditionalFormattingErr shtGlobGuiMedDisc, "R", strConcErr, 9, 38, 7
    SetConditionalFormattingErr shtGlobGuiMedDisc, "S", strTimeErr, 9, 38, 7
    
    SetConditionalFormattingInfo shtGlobGuiMedDisc, "F", strMOWarn, 9, 38, 7

End Sub

Public Sub SetMedDiscPrtConditionalFormatting()
    
    Dim strFreqWarn As String
    Dim strDoseWarn As String
    Dim strDoseErr As String
    
    Dim strConcErr As String
    Dim strOplErr As String
    Dim strTimeErr As String
    
    Dim strMOWarn As String
    
    strFreqWarn = "=GlobBerMedDisc!BF"
    strDoseWarn = "=GlobBerMedDisc!BL"
    strDoseErr = "=GlobBerMedDisc!BM"
    
    strConcErr = "=GlobBerMedDisc!BA"
    strOplErr = "=GlobBerMedDisc!BB"
    strTimeErr = "=GlobBerMedDisc!BC"
        
    SetConditionalFormattingWarn shtGlobPrtMedDisc, "B", strFreqWarn, 3, 32, 1
    SetConditionalFormattingWarnErr shtGlobPrtMedDisc, "F", strDoseErr, strDoseWarn, 3, 32, 1
    
    SetConditionalFormattingErr shtGlobPrtMedDisc, "E", strConcErr, 3, 32, 1
    SetConditionalFormattingErr shtGlobPrtMedDisc, "E", strTimeErr, 3, 32, 1
    
End Sub
Public Sub SetInfBVochtTPNFormatting()
    
    Dim strTPNError As String
    
    strTPNError = "=NeoBerInfB!AE"
    
    SetConditionalFormattingErr shtNeoGuiInfB, "G", strTPNError, 59, 59, 9

End Sub

Public Sub SetInfB1700Formatting()
    
    Dim strTPNError As String
    
    strTPNError = "=NeoBerInfB!AE"
    
    SetConditionalFormattingErr shtNeoGuiInfB, "C", strTPNError, 9, 9, 6

End Sub

Public Sub SetPedTPNNegLipidVolFormatting()
    
    Dim strTPNError As String
    
    strTPNError = "=PedBerTPN!D"
    ' 17 - 35
    SetConditionalFormattingErr shtPedGuiEntTPN, "G", strTPNError, 35, 35, 18

End Sub

Public Sub SetPedTPNSST1LT24Formatting()
    
    Dim strTPNError As String
    
    strTPNError = "=PedBerTPN!AE"
    ' 3 - 22
    SetConditionalFormattingErr shtPedGuiEntTPN, "I", strTPNError, 22, 22, 19

End Sub

Public Sub SetPedTPNLipidLT24Formatting()
    
    Dim strTPNError As String
    
    strTPNError = "=PedBerTPN!AE"
    ' 16 - 34
    SetConditionalFormattingErr shtPedGuiEntTPN, "I", strTPNError, 34, 34, 18

End Sub


