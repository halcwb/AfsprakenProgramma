Attribute VB_Name = "ModOpenClose"
Option Explicit

Private intCount As Integer

Public blnDontClose As Boolean

Public Sub SetToDevelopmentMode()

    Dim objSheet As Worksheet
    Dim colSheets As Collection
    
    blnDontClose = True
        
    InterfaceSheetsUnprotect
    UnhideNonUserInterfaceSheets
    
    Application.DisplayFormulaBar = True
    
    Set colSheets = GetNonInterfaceSheets()
    For Each objSheet In colSheets
        objSheet.Activate
        SetWindowToClose Windows(1)
    Next
    
    Set colSheets = GetUserInterfaceSheets()
    For Each objSheet In colSheets
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
    LogActionStart strAction, strParams
    
    Dim objWindow As Window

    Application.Cursor = xlWait
    Application.DisplayAlerts = False
    
    For Each objWindow In Application.Windows
        SetWindowToClose objWindow
    Next
 
    Toolbars("Afspraken").visible = False
    
    With Application
         .Caption = vbNullString
         .DisplayFormulaBar = True
         .Cursor = xlDefault
         If Not blnDontClose Then .Quit
    End With
    
    LogActionEnd "Afsluiten"
            
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
    LogActionStart strAction, strParams
    
    Dim objWindow As Window
    
    Application.Cursor = xlWait
    Workbooks(CONST_WORKBOOKNAME).Activate

    ProtectUserInterfaceSheets
    HideAndUnProtectNonUserInterfaceSheets

'   Knoppen en balken verwijderen
    HideBars
    ActiveWindow.DisplayWorkbookTabs = BlnIsDevelopment

    For Each objWindow In Application.Windows
        SetWindowToOpen objWindow
    Next
    
'   Zorg ervoor dat niet per ongeluk een lege patient naar een bed wordt weggeschreven
    Range("BedNummer").Value = 0
    Range("AfspraakDatum").FormulaLocal = "=Vandaag()"
    
'   verwijder afspraken
    shtGuiAcuteOpvang.Unprotect CONST_PASSWORD 'ICT2014
    
    With shtPatData
            
        For intCount = 2 To .Range("A1").CurrentRegion.Rows.Count
            Range(.Cells(intCount, 1)).Formula = .Cells(intCount, 3).Formula
        Next intCount
    
    End With
    
    Range("AfsprakenVersie").Value = vbNullString
    
    shtGuiAcuteOpvang.Protect CONST_PASSWORD 'ICT2014
    shtGuiAcuteOpvang.Activate
    
    BlnEnableDevelop = False
    
    Application.Cursor = xlDefault
    
    LogActionEnd "Openen"

End Sub

Private Sub InterfaceSheetsUnprotect()
            
    Dim col As New Collection, intCount As Integer
        
    Set col = GetUserInterfaceSheets()
    
    For intCount = 1 To col.Count
    
        With col(intCount)
            .EnableSelection = xlNoRestrictions
            .Unprotect PASSWORD:=CONST_PASSWORD
       End With
       
    Next intCount

    Set col = Nothing

End Sub

Private Sub ProtectUserInterfaceSheets()
            
    Dim col As New Collection, intCount As Integer
        
    Set col = GetUserInterfaceSheets()
    
    For intCount = 1 To col.Count
    
        With col(intCount)
            .EnableSelection = xlNoSelection
            .Protect PASSWORD:=CONST_PASSWORD
            .DisplayPageBreaks = False
       End With
       
    Next intCount

    Set col = Nothing

End Sub

' Get all sheets that act as a User Interface
' Must be visible and protected
Public Function GetUserInterfaceSheets() As Collection
'TODO: Update list of Interface sheets
    Dim col As New Collection
        
    With col
        .Add Item:=shtGuiAcuteOpvang
        .Add Item:=shtGuiInfusen
        .Add Item:=shtGuiIntake
        .Add Item:=shtGuiLab
        .Add Item:=shtGuiMedDisc
        .Add Item:=shtGuiMedicatieIV
        .Add Item:=shtPrtAfspraken
        .Add Item:=shtPrtMedicatie
        .Add Item:=shtPrtTPN16tot30kg
        .Add Item:=shtPrtTPN2tot6kg
        .Add Item:=shtPrtTPN31tot50kg
        .Add Item:=shtPrtTPN7tot15kg
        .Add Item:=shtPrtTPNboven50kg
    
        .Add Item:=shtAanvullendeAfspraken
        .Add Item:=shtAanvullendeAfsprakenPed
        .Add Item:=shtGuiAfspraken
        .Add Item:=shtGuiAfspraken1700
        .Add Item:=shtApotheek
        .Add Item:=shtGuiAcuteOpvangNeo
        .Add Item:=shtGuiLabNeo
        .Add Item:=shtPrint
        .Add Item:=shtGuiWerkBrief
    
    End With
    
    Set GetUserInterfaceSheets = col

End Function

' Get all sheets that do work and are not User Interface
' Must be hidden and not protected
Public Function GetNonInterfaceSheets() As Collection
'TODO: Update list of Calculation sheets

    Dim col As New Collection
    
    With col
        .Add Item:=shtBerConversie
        .Add Item:=shtBerInfusen
        .Add Item:=shtBerIVMed
        .Add Item:=shtBerLab
        .Add Item:=shtBerMedDisc
        .Add Item:=shtBerNormaal
        .Add Item:=shtBerOpm
        .Add Item:=shtBerPO
        .Add Item:=shtBerTemp
        .Add Item:=shtBerTijden
        .Add Item:=shtBerTotalen
        .Add Item:=shtBerTPN
        .Add Item:=shtPatAfsprakenTekst
        .Add Item:=shtPatDetails
        .Add Item:=shtPatData
        .Add Item:=shtTblHeightNL
        .Add Item:=shtTblWeigthNL
        .Add Item:=shtBerTijdenNeo
        .Add Item:=shtTblMedDisc
        .Add Item:=shtTblInfusen
    
        .Add Item:=shtTblMedicatieIV
        .Add Item:=shtTblMedicatieIVNeo
        .Add Item:=shtTblTPOSheet1
        .Add Item:=shtBerekeningen
        .Add Item:=shtBerekeningen1700
        .Add Item:=shtBerLabNeo
        .Add Item:=shtAanvullendeTbl
        .Add Item:=shtAanvullendeBer
        .Add Item:=shtAanvullendeBerPed
        .Add Item:=shtAdvies
        .Add Item:=shtLijsten
        .Add Item:=shtTblTPOSheet1
        .Add Item:=shtTblVoeding
        
        .Add Item:=shtDivPediatrie
        .Add Item:=shtDivNeo
        .Add Item:=shtDivPatient
        .Add Item:=shtWerkBriefActueel
        
        
    End With

    Set GetNonInterfaceSheets = col

End Function

Private Sub HideAndUnProtectNonUserInterfaceSheets()

    Dim col As New Collection
    
    Set col = GetNonInterfaceSheets()
    
    For intCount = 1 To col.Count
        With col(intCount)
            .visible = xlVeryHidden
            .Unprotect PASSWORD:=CONST_PASSWORD
        End With
    Next intCount

    Set col = Nothing

End Sub

Private Sub UnhideNonUserInterfaceSheets()

    Dim col As New Collection
    
    Set col = GetNonInterfaceSheets()
    
    For intCount = 1 To col.Count
        With col(intCount)
            .visible = True
        End With
    Next intCount

    Set col = Nothing

End Sub

Sub HideBars()
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
' If peli or developper then shtGUIMedicatieIV
' Else shtAfspraken
Sub OpenStartSheet()

    Dim strPath As String
    Dim strPeli As String

    SetDeveloperMode
    
    strPath = Application.ActiveWorkbook.Path
    strPeli = ModGlobal.CONST_PELI_FOLDERNAME
    
    If ModString.StringContainsCaseInsensitive(strPath, strPeli) Or BlnIsDevelopment Then
        shtGuiMedicatieIV.Select
    Else
        shtGuiAfspraken.Select
    End If
    
End Sub

Sub OpenLabSheet()
Dim strPath As String

    SetDeveloperMode
    
    strPath = LCase(Application.ActiveWorkbook.Path)
    If InStr(1, strPath, LCase(CONST_PELI_FOLDERNAME)) > 0 Or BlnIsDevelopment > 0 Then
        shtGuiLab.Select
    Else
        shtGuiLabNeo.Select
    End If
End Sub

Sub OpenAanvullendeSheet()
Dim strPath As String

    SetDeveloperMode
    
    strPath = LCase(Application.ActiveWorkbook.Path)
    If InStr(1, strPath, LCase(CONST_PELI_FOLDERNAME)) > 0 Or BlnIsDevelopment > 0 Then
        shtAanvullendeAfsprakenPed.Select
    Else
        shtAanvullendeAfspraken.Select
    End If
End Sub


