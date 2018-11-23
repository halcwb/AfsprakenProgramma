Attribute VB_Name = "ModSheet"
Option Explicit

Private m_blnPrev As Boolean

Public Function IsUserInterface(shtSheet As Worksheet) As Boolean

    Dim blnGui As Boolean
    
    blnGui = ModString.ContainsCaseSensitive(shtSheet.Name, "Gui")
    blnGui = blnGui Or ModString.ContainsCaseSensitive(shtSheet.Name, "Prt")
    
    IsUserInterface = blnGui

End Function

' Get all sheets that act as a User Interface
' Must be visible and protected
Public Function GetUserInterfaceSheets() As Collection
    
    Dim colShts As Collection
    Dim shtSheet As Worksheet
    
    Set colShts = New Collection
    For Each shtSheet In WbkAfspraken.Sheets
    
        If IsUserInterface(shtSheet) Then
            colShts.Add shtSheet
        End If
    
    Next shtSheet
    
    Set GetUserInterfaceSheets = colShts

End Function

' Get all sheets that do work and are not User Interface
' Must be hidden and not protected
Public Function GetNonInterfaceSheets() As Collection
    
    Dim colShts As Collection
    Dim shtSheet As Worksheet
    
    Set colShts = New Collection
    For Each shtSheet In WbkAfspraken.Sheets
    
        If Not IsUserInterface(shtSheet) Then
            colShts.Add shtSheet
        End If
    
    Next shtSheet
    
    Set GetNonInterfaceSheets = colShts

End Function

Public Sub UnhideNonUserInterfaceSheets(ByVal blnShowProgress As Boolean)

    Dim colShts As Collection
    Dim shtSheet As Worksheet
    Dim intN As Integer
    Dim intC As Integer
    
    Set colShts = GetNonInterfaceSheets()
    intN = 1
    intC = colShts.Count
    
    For Each shtSheet In colShts
    
        shtSheet.Visible = True
        
        If blnShowProgress Then ModProgress.SetJobPercentage "Verberg Bladen", intC, intN
        intN = intN + 1

    Next shtSheet

End Sub

Public Sub HideAndUnProtectNonUserInterfaceSheets(ByVal blnShowProgress As Boolean)
 
    Dim colShts As Collection
    Dim shtSheet As Worksheet
    Dim intN As Integer
    Dim intC As Integer
    
    Set colShts = GetNonInterfaceSheets()
    intN = 1
    intC = colShts.Count
    
    For Each shtSheet In colShts
    
        With shtSheet
            .Visible = xlVeryHidden
            .Unprotect Password:=CONST_PASSWORD
        End With
        
        If blnShowProgress Then ModProgress.SetJobPercentage "Verberg reken bladen ...", intC, intN
        intN = intN + 1

    Next shtSheet

End Sub

Public Sub UnprotectUserInterfaceSheets(ByVal blnShowProgress As Boolean)

    Dim objItem As Worksheet
    Dim colShts As Collection
    Dim intN As Integer
    Dim intC As Integer
    
    Set colShts = GetUserInterfaceSheets()
    intN = 0
    intC = colShts.Count
    For Each objItem In colShts
    
        With objItem
            .Visible = xlSheetVisible
            .EnableSelection = xlNoRestrictions
            .DisplayPageBreaks = True
            .Unprotect ModConst.CONST_PASSWORD
        End With
        
        intN = intN + 1
        If blnShowProgress Then ModProgress.SetJobPercentage "Verwijder Beveiliging", intC, intN
        
    Next objItem
            
End Sub

Public Sub ProtectUserInterfaceSheets(ByVal blnShowProgress As Boolean)
            
    Dim objItem As Worksheet
    Dim colShts As Collection
    Dim intN As Integer
    Dim intC As Integer
    
    Set colShts = GetUserInterfaceSheets()
    intN = 0
    intC = colShts.Count
    For Each objItem In colShts
    
        With objItem
            .Visible = xlSheetVisible
            .EnableSelection = xlNoSelection
            .DisplayPageBreaks = False
            .Protect ModConst.CONST_PASSWORD
        End With
        
        intN = intN + 1
        If blnShowProgress Then ModProgress.SetJobPercentage "Stel beveiliging in", intC, intN
        
    Next objItem

End Sub

Public Sub GoToSheet(shtSheet As Worksheet, ByVal strRange As String)

    shtSheet.Select
    shtSheet.Range(strRange).Select
    ActiveWindow.ScrollRow = 1

End Sub

Public Function GetInterfaceSheetCount() As Integer

    Dim intN As Integer
    Dim shtSheet As Worksheet
    
    For Each shtSheet In ActiveWorkbook.Sheets
    
        If IsUserInterface(shtSheet) Then
            intN = intN + 1
        End If
    
    Next shtSheet
    
    GetInterfaceSheetCount = intN

End Function

Public Function GetNonInterfaceSheetCount() As Integer

    Dim intN As Integer
    
    intN = ActiveWorkbook.Sheets.Count
    GetNonInterfaceSheetCount = intN - GetInterfaceSheetCount()

End Function

Public Sub SelectNeoOrPedSheet(shtPed As Worksheet, shtNeo As Worksheet, ByVal blnStartProgress As Boolean)

    If MetaVision_IsPediatrie() Then
        GoToSheet shtPed, "A1"
    Else
        If shtNeo.Name = shtNeoGuiInfB.Name Then
            ModNeoInfB.NeoInfB_SelectInfB False, blnStartProgress
        Else
            GoToSheet shtNeo, "A1"
        End If
    End If
    
End Sub

Public Sub SelectPedOrNeoStartSheet(ByVal blnStartProgress As Boolean)

    SelectNeoOrPedSheet shtPedGuiMedIV, shtNeoGuiInfB, blnStartProgress
    
End Sub

Private Sub Test_SelectPedOrNeoStartSheet()

    SelectPedOrNeoStartSheet True

End Sub


Public Sub SelectPedOrNeoLabSheet()
    
    SelectNeoOrPedSheet shtPedGuiLab, shtNeoGuiLab, False
        
End Sub

Public Sub SelectPedOrNeoAfsprSheet()
    
    SelectNeoOrPedSheet shtPedGuiAfspr, shtNeoGuiAfspr, False
        
End Sub

Public Sub PrintSheet(shtSheet As Worksheet, ByVal intNum As Integer, ByVal blnAsk As Boolean, ByVal blnPrev As Boolean)
  
    shtSheet.Unprotect ModConst.CONST_PASSWORD
    shtSheet.PageSetup.LeftHeader = "AfsprakenProgramma " & ModRange.GetRangeValue("Var_Glob_AppVersie", vbNullString)
    
    If blnAsk Then
        m_blnPrev = ModMessage.ShowMsgBoxYesNo("Eerst preview zien?") = vbYes
    Else
        m_blnPrev = blnPrev
    End If
    
    If m_blnPrev Then
        shtSheet.PrintPreview False
    Else
        shtSheet.PrintOut Copies:=intNum
    End If
    
    If Not ModSetting.IsDevelopmentMode Then shtSheet.Protect ModConst.CONST_PASSWORD
    
End Sub

Public Sub SaveSheetAsPDF(shtSheet As Worksheet, ByVal strFile As String, blnFitToPortrait As Boolean)
    
    shtSheet.Unprotect ModConst.CONST_PASSWORD
    shtSheet.PageSetup.LeftHeader = "AfsprakenProgramma " & ModRange.GetRangeValue("Var_Glob_AppVersie", vbNullString)

    If blnFitToPortrait Then
        With shtSheet.PageSetup
            .Orientation = xlPortrait
            .Draft = False
            .FirstPageNumber = xlAutomatic
            .Order = xlDownThenOver
            .BlackAndWhite = False
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
        End With
    End If
        
    shtSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=strFile, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False

    If Not ModSetting.IsDevelopmentMode Then shtSheet.Protect ModConst.CONST_PASSWORD

End Sub

Public Sub PrintSheetAllLandscape(shtSheet As Worksheet)
    
    With shtSheet.PageSetup
        .Orientation = xlLandscape
        .Draft = False
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With

End Sub

Public Sub PrintSheetAllPortrait(shtSheet As Worksheet)
    
    With shtSheet.PageSetup
        .Orientation = xlPortrait
        .Draft = False
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With

End Sub

Public Sub Sheet_CopyRangeFormulaToDst(objSrc As Range, objDst As Range)
           
    On Error GoTo CopyRangeFormulaToDstError
    
    If Not objSrc.Rows.Count = objDst.Rows.Count Or Not objSrc.Columns.Count = objDst.Columns.Count Then Err.Raise ModConst.CONST_APP_ERROR, , ModConst.CONST_DEFAULTERROR_MSG
    
    objSrc.Copy
    
    objDst.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
    Exit Sub

CopyRangeFormulaToDstError:

    ModLog.LogError Err, "Cannot copy " & objSrc.Address & " to " & objDst.Address
    Err.Raise ModConst.CONST_APP_ERROR, , ModConst.CONST_DEFAULTERROR_MSG

End Sub



