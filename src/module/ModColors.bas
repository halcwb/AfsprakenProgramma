Attribute VB_Name = "ModColors"
Option Explicit

Public Enum RGBColors
    R = 1
    G = 2
    B = 3
End Enum

Private Const constColorSettings As String = "G2"
Private Const constPedTable As String = "K2"

Private Sub SetRangeColor(ByRef objTarget As Range, objSetting As Range, ByVal blnSheet As Boolean, Optional ByVal varGridColor As Variant)

    Dim lngGridColor As Long
    
    If blnSheet Then
        Set objTarget = objTarget.Worksheet.Cells
    Else
        objTarget.Font.Color = objSetting.Font.Color
        objTarget.Font.Name = objSetting.Font.Name
        objTarget.Font.Bold = objSetting.Font.Bold
        objTarget.Font.Italic = objSetting.Font.Italic
    End If

    objTarget.Interior.Color = objSetting.Interior.Color
    
    If Not IsMissing(varGridColor) Then
        lngGridColor = Conversion.CLng(varGridColor)
        
        objTarget.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        objTarget.Borders(xlInsideHorizontal).Weight = xlThick
        objTarget.Borders(xlInsideHorizontal).Color = lngGridColor
    
        objTarget.Borders(xlInsideVertical).LineStyle = xlContinuous
        objTarget.Borders(xlInsideVertical).Weight = xlThick
        objTarget.Borders(xlInsideVertical).Color = lngGridColor
    End If

End Sub

Public Sub ColorPedRanges()

    Dim objPed As Range
    Dim objSettings As Range
    Dim intN As Integer
    Dim intC As Integer
    
    Dim strSetting As String
    Dim objSetting As Range
    Dim blnSheet As Boolean
    
    Dim intTargetN As Integer
    Dim intTargetC As Integer
    Dim strSheet As String
    Dim strTarget As String
    Dim strRange As String
    Dim objTarget As Range
    
    Dim lngBackGround As Long
    
    Set objSettings = shtGlobSettings.Range(constColorSettings).CurrentRegion
    Set objPed = shtGlobSettings.Range(constPedTable).CurrentRegion
    
    ModProgress.StartProgress "Kleuren Instellen"
    
    intC = objSettings.Rows.Count
    For intN = 2 To intC
        strSetting = objSettings.Cells(intN, 1).Value2
        blnSheet = False
        
        ModProgress.SetJobPercentage strSetting, intC, intN
        
        If strSetting = "Backgrounds" Then
            lngBackGround = objSettings.Cells(intN, 2).Interior.Color
            blnSheet = True
        End If
        
        intTargetC = objPed.Rows.Count
        For intTargetN = 2 To intTargetC
            
            strSheet = objPed.Cells(intTargetN, 1).Value2
            strTarget = objPed.Cells(intTargetN, 2).Value2
            strRange = Strings.Replace(objPed.Cells(intTargetN, 3).Formula, "=", vbNullString)
            strRange = Strings.Replace(strRange, strSheet & "!", vbNullString)
            
            If strTarget = strSetting And Not strRange = vbNullString Then
                Set objSetting = objSettings.Cells(intN, 2)
                Set objTarget = WbkAfspraken.Sheets(strSheet).Range(strRange)
                
                If strSetting = "Fields" Then
                    SetRangeColor objTarget, objSetting, blnSheet, lngBackGround
                Else
                    SetRangeColor objTarget, objSetting, blnSheet
                End If
                
                Set objSetting = Nothing
                Set objTarget = Nothing
            End If
        Next
    Next
    
    ModProgress.FinishProgress

End Sub

Public Function ConvertLongToRGB(ByVal lngC As Long, ByVal intOpt As RGBColors) As Long

    Dim lngR As Long
    Dim lngG As Long
    Dim lngB As Long

    lngR = lngC Mod 256
    lngG = lngC \ 256 Mod 256
    lngB = lngC \ 65536 Mod 256

    If intOpt = RGBColors.R Then
        ConvertLongToRGB = lngR
    ElseIf intOpt = RGBColors.G Then
        ConvertLongToRGB = lngG
    Else
        ConvertLongToRGB = lngB
    End If

End Function

Public Function ShowColorDialog(ByVal lngC As Long) As Long

    Dim lngR As Long
    Dim lngG As Long
    Dim lngB As Long

    lngR = ConvertLongToRGB(lngC, R)
    lngG = ConvertLongToRGB(lngC, G)
    lngB = ConvertLongToRGB(lngC, B)

    If Application.Dialogs(xlDialogEditColor).Show(10, lngR, lngG, lngB) = True Then
      'user pressed OK
      ShowColorDialog = ActiveWorkbook.Colors(10)
    Else
      'user pressed Cancel
      ShowColorDialog = -1
    End If

End Function

Public Function ShowFontDialog(ByRef objRange As Range) As Boolean

    objRange.Select
        
    If Application.Dialogs(xlDialogFontProperties).Show Then
        'User accepted, check what has changed
        ShowFontDialog = True
    Else
        'User cancelled
        ShowFontDialog = False
    End If

End Function

Public Function GetFontNames() As Variant()

    Dim objWd As Object
    Dim varID As Variant
    Dim varIds() As Variant

    Set objWd = CreateObject("Word.Application")

    For Each varID In objWd.FontNames
        ModArray.AddItemToVariantArray varIds, varID
    Next
    
    objWd.Quit
    Set objWd = Nothing
    GetFontNames = varIds

End Function

Private Sub TestGetFontNames()

    Dim varN As Variant
    
    For Each varN In GetFontNames
        MsgBox varN
    Next

End Sub

Private Sub TestShowColorDialog()

    MsgBox ShowColorDialog(979)

End Sub
