Attribute VB_Name = "ModOpenClose"
Option Explicit

Private intCount As Integer

Public blnDontClose As Boolean

Public Sub SetToDevelopmentMode()

    Dim objSheet As Worksheet
    
    blnDontClose = True
        
    ModSheets.UnprotectUserInterfaceSheets
    ModSheets.UnhideNonUserInterfaceSheets
    
    Application.DisplayFormulaBar = True
    
    For Each objSheet In ModSheets.GetNonInterfaceSheets()
        objSheet.Activate
        SetWindowToClose Windows(1)
    Next
    
    For Each objSheet In ModSheets.GetUserInterfaceSheets()
        objSheet.Activate
        SetWindowToClose Windows(1)
    Next
    
    blnDontClose = False
    BlnEnableDevelop = True ' BlnIsDevelopment
    
    Application.Cursor = xlDefault

End Sub

Public Sub Afsluiten()

    Dim strAction As String, strParams() As Variant
    strAction = "Afsluiten"
    strParams = Array()
    
    ModLogging.LogActionStart strAction, strParams
    
    Dim objWindow As Window

    Application.Cursor = xlWait
    Application.DisplayAlerts = False
    
    For Each objWindow In Application.Windows
        SetWindowToClose objWindow
    Next
 
    Toolbars("Afspraken").Visible = False
    
    With Application
         .Caption = vbNullString
         .DisplayFormulaBar = True
         .Cursor = xlDefault
         If Not blnDontClose Then .Quit
    End With
    
    ModLogging.LogActionEnd "Afsluiten"
            
End Sub

Private Sub SetWindow(objWindow As Window, blnReset As Boolean)

    With objWindow
        If BlnIsDevelopment Then
            .DisplayWorkbookTabs = True
            .DisplayGridlines = True
            .DisplayHeadings = True
            .DisplayOutline = True
            .DisplayZeros = True
        Else
            .DisplayGridlines = blnReset
            .DisplayHeadings = blnReset
            .DisplayOutline = blnReset
            .DisplayZeros = blnReset
            .DisplayWorkbookTabs = blnReset
        End If
    End With

End Sub

Public Sub SetWindowToClose(objWindow As Window)
    
    SetWindow objWindow, True

End Sub

Public Sub SetWindowToOpen(objWindow As Window)
    
    SetWindow objWindow, False

End Sub

Sub Openen()
Attribute Openen.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim strAction As String, strParams() As Variant
    strAction = "Openen"
    strParams = Array()
    
    ModLogging.LogActionStart strAction, strParams
    
    Dim objWindow As Window
    
    Application.Cursor = xlWait
    Workbooks(CONST_WORKBOOKNAME).Activate

    ModSheets.ProtectUserInterfaceSheets
    ModSheets.HideAndUnProtectNonUserInterfaceSheets

'   Knoppen en balken verwijderen
    SetCaptionHideBars
    ActiveWindow.DisplayWorkbookTabs = BlnIsDevelopment

    For Each objWindow In Application.Windows
        SetWindowToOpen objWindow
    Next
    
'   Zorg ervoor dat niet per ongeluk een lege patient naar een bed wordt weggeschreven
    Range("BedNummer").Value = 0
    Range("AfspraakDatum").FormulaLocal = "=Vandaag()"
    
'   verwijder afspraken
    shtPedGuiAcuut.Unprotect CONST_PASSWORD 'ICT2014
    
    With shtPatData
            
        For intCount = 2 To .Range("A1").CurrentRegion.Rows.Count
            Range(.Cells(intCount, 1)).Formula = .Cells(intCount, 3).Formula
        Next intCount
    
    End With
    
    Range("AfsprakenVersie").Value = vbNullString
    
    shtPedGuiAcuut.Protect CONST_PASSWORD 'ICT2014
    shtPedGuiAcuut.Activate
    
    BlnEnableDevelop = False
    
    Application.Cursor = xlDefault
    
    ModLogging.LogActionEnd "Openen"

End Sub

Private Sub SetCaptionHideBars()
'   Knoppen en balken verwijderen
    With Application
         .Caption = "Informedica 2015 Afspraken programma"
         .DisplayFormulaBar = BlnIsDevelopment
         .DisplayStatusBar = BlnIsDevelopment
         .DisplayScrollBars = True
         .DisplayFormulaBar = BlnIsDevelopment
    End With
    
End Sub

' Determine the sheet to open with
' If peli or developper then ped sheet
' Else neo sheet
Public Sub SelectNeoOrPedSheet(shtPed As Worksheet, shtNeo As Worksheet)

    Dim strPath As String
    Dim strPeli As String

    ModConst.SetDeveloperMode
    
    strPath = Application.ActiveWorkbook.Path
    strPeli = ModConst.CONST_PELI_FOLDERNAME
    
    If ModString.StringContainsCaseInsensitive(strPath, strPeli) Or BlnIsDevelopment Then
        shtPed.Select
    Else
        shtNeo.Select
    End If
    
End Sub

Public Sub SelectPedOrNeoStartSheet()

    SelectNeoOrPedSheet shtPedGuiMedIV, shtNeoGuiAfspraken
    
End Sub

Public Sub SelectPedOrNeoLabSheet()
    
    SelectNeoOrPedSheet shtPedGuiLab, shtNeoGuiLab
        
End Sub

Public Sub SelectPedOrNeoAfsprExtraSheet()
    
    SelectNeoOrPedSheet shtPedGuiAfsprExta, shtNeoGuiAfsprExtra
        
End Sub


