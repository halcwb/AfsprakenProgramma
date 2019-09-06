Attribute VB_Name = "ModColors"
Option Explicit

Public Enum RGBColors
    R = 1
    g = 2
    B = 3
End Enum

Private Const constColorSettings As String = "G2"
Private Const constSheetRangeTable As String = "K2"

Private Sub ClearSheetTableBorders(objSheet As Worksheet)
    
    With objSheet.Cells
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
    
End Sub


Private Function IsNeoSheet(ByVal strSheet As String) As Boolean

    IsNeoSheet = ModString.StartsWith(strSheet, "NeoGui")

End Function

Private Sub NoFill(objTarget As Range)

    With objTarget.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

End Sub

' ToDo Grid lines thick is somehow not working for Neo
Private Sub SetRangeColor(objTarget As Range, objSetting As Range, ByVal blnSheet As Boolean, Optional ByVal varGridColor As Variant)

    Dim lngGridColor As Long
    Dim objCell As Range
    
    If blnSheet Then
        Set objTarget = objTarget.Worksheet.Cells
    Else
        objTarget.Font.Color = objSetting.Font.Color
        objTarget.Font.Name = objSetting.Font.Name
        objTarget.Font.Bold = objSetting.Font.Bold
        objTarget.Font.Italic = objSetting.Font.Italic
    End If

    objTarget.Interior.Color = objSetting.Interior.Color
    
    If Not IsMissing(varGridColor) And objTarget.Rows.Count > 1 Then
        lngGridColor = Conversion.CLng(varGridColor)
                    
        For Each objCell In objTarget.Cells
            objCell.Borders(xlEdgeBottom).LineStyle = xlContinuous
            objCell.Borders(xlEdgeBottom).Weight = xlThick
            objCell.Borders(xlEdgeBottom).Color = lngGridColor
        Next
        
    End If

End Sub

Private Sub Test_RowsCount()

    MsgBox shtPedGuiEntTPN.Range("D14:E16").Rows.Count

End Sub

Public Sub ColorPedNeoRanges(ByVal blnNeo As Boolean)

    Dim objSheetRanges As Range
    Dim objSettings As Range
    Dim intN As Integer
    Dim intC As Integer
    
    Dim strSetting As String
    Dim objSetting As Range
    Dim intSetting As Integer
    Dim blnSheet As Boolean
    
    Dim intTargetN As Integer
    Dim intTargetC As Integer
    Dim strSheet As String
    Dim strTarget As String
    Dim strRange As String
    Dim varRangeItem As Variant
    Dim objTargetSheet As Worksheet
    Dim objTargetRange As Range
    
    Dim lngBackGround As Long
    Dim blnProtected As Boolean
       
    Set objSettings = shtGlobSettings.Range(constColorSettings).CurrentRegion
    Set objSheetRanges = shtGlobSettings.Range(constSheetRangeTable).CurrentRegion
    
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
        
        intTargetC = objSheetRanges.Rows.Count
        For intTargetN = 2 To intTargetC
            
            strSheet = objSheetRanges.Cells(intTargetN, 1).Value2
            
            If blnNeo And IsNeoSheet(strSheet) Then
                intSetting = 3
            ElseIf Not IsNeoSheet(strSheet) Then
                intSetting = 2
            Else
                intSetting = -1
            End If
            
            strTarget = objSheetRanges.Cells(intTargetN, 2).Value2
            strRange = Strings.Replace(objSheetRanges.Cells(intTargetN, 3).Formula, "=", vbNullString)
            strRange = Strings.Replace(strRange, strSheet & "!", vbNullString)
            
            blnProtected = False
            If strTarget = strSetting And Not strRange = vbNullString And Not intSetting = -1 Then
                ModLog.LogInfo "Coloring: " & Join(Array(strSheet, strSetting, strRange), " ")
                
                Set objSetting = objSettings.Cells(intN, intSetting)
                Set objTargetSheet = WbkAfspraken.Sheets(strSheet)
                
                If objTargetSheet.ProtectContents Then
                    blnProtected = True
                    objTargetSheet.Unprotect ModConst.CONST_PASSWORD
                End If
                
                If strSetting = "Backgrounds" Then ClearSheetTableBorders objTargetSheet
                
                For Each varRangeItem In Split(strRange, ",")
                
                    Set objTargetRange = objTargetSheet.Range(CStr(varRangeItem))
                    
                    If strSetting = "Fields" Then
                        SetRangeColor objTargetRange, objSetting, blnSheet, lngBackGround
                    Else
                        SetRangeColor objTargetRange, objSetting, blnSheet
                    End If
                
                Next varRangeItem
                
                If blnProtected Then objTargetSheet.Protect ModConst.CONST_PASSWORD
                
            End If
            
            ModProgress.SetJobPercentage strSheet & ": " & strTarget, intTargetC, intTargetN
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
    ElseIf intOpt = RGBColors.g Then
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
    lngG = ConvertLongToRGB(lngC, g)
    lngB = ConvertLongToRGB(lngC, B)

    If Application.Dialogs(xlDialogEditColor).Show(10, lngR, lngG, lngB) = True Then
      'user pressed OK
      ShowColorDialog = ActiveWorkbook.Colors(10)
    Else
      'user pressed Cancel
      ShowColorDialog = -1
    End If

End Function

Public Function ShowFontDialog(objRange As Range) As Boolean

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

Private Sub Test_GetFontNames()

    Dim varN As Variant
    
    For Each varN In GetFontNames
        MsgBox varN
    Next

End Sub

Private Sub Test_ShowColorDialog()

    MsgBox ShowColorDialog(979)

End Sub
