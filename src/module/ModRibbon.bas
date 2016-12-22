Attribute VB_Name = "ModRibbon"
Option Explicit

Public Sub ButtonOnAction(ctrlMenuItem As IRibbonControl)
'
' Code for onAction callback. Ribbon control button
'

    Application.ScreenUpdating = False
    
    Select Case ctrlMenuItem.ID
        'GroupAfspraken
        Case "btnAfsluiten"
            Afsluiten
        Case "btnAfsprakenVerwijderen"
            ClearPatient True
            SelectPedOrNeoStartSheet
        'GroupBedden
        Case "btnBedOpenen"
            OpenPatientLijst
        Case "btnBedOpslaan"
            ModBedden.SluitBed
            SelectPedOrNeoStartSheet
        Case "btnGegevensInvoeren"
            NieuwePatient
        'grpPediatrie
        Case "btnPContinueivmedicatie"
            ModSheets.GoToSheet shtPedGuiMedIV, "A6"
        Case "btnPDiscontinuemedicatie"
            ModSheets.GoToSheet shtPedGuiMedDisc, "A6"
        Case "btnPInfusen"
            ModSheets.GoToSheet shtPedGuiPMenIV, "A6"
        Case "btnPIntake"
            ModSheets.GoToSheet shtPedGuiEntTPN, "A6"
        Case "btnPLaboratoriumbepalingen"
            ModSheets.GoToSheet shtPedGuiLab, "A1"
        Case "btnPAanvullendeAfspraken"
            ModSheets.GoToSheet shtPedGuiAfsprExta, "A6"
        'grpNeonatologie
        Case "btnNInfuusbrief"
            ModSheets.GoToSheet shtNeoGuiAfspraken, "A9"
        Case "btnNDiscontinuemedicatie"
            ModSheets.GoToSheet shtPedGuiMedDisc, "A6"
        Case "btnNAanvullendeAfspraken"
            ModSheets.GoToSheet shtNeoGuiAfsprExtra, "A6"
        Case "btnNTPNadvies"
            TPNAdviesNEO
        Case "btnNLaboratoriumbepalingen"
            ModSheets.GoToSheet shtNeoGuiLab, "A1"
        Case "btnNAfspraken1700"
            ModSheets.GoToSheet shtNeoGuiAfspr1700, "A9"
        'grpGegevensVerwijderen -> Acties
        Case "btnNAfspraken1700Overzetten"
            ModAfspraken1700.CopyToActueel
        Case "btnNActueleAfsprakenOverzetten"
            ModAfspraken.AfsprakenOvernemen
        Case "btnLabVerwijderen"
            VerwijderLab
            SelectPedOrNeoLabSheet
        Case "btnAanvullendVerwijderen"
            VerwijderAanvullendeAfspraken
            SelectPedOrNeoAfsprExtraSheet
        'grpPrintPediatrie
        Case "btnPAcuteBlad"
            ModSheets.GoToSheet shtPedGuiAcuut, "A1"
        Case "btnPPrintAfspraken"
            ModSheets.GoToSheet shtPedPrtAfspr, "A1"
        Case "btnPMedicatie"
            ModSheets.GoToSheet shtPedPrtMedDisc, "A1"
        Case "btnPTPN"
            SelectPedTPNPrint
        'grpPrintNeonatologie
        Case "btnNAcuteBlad"
            ModSheets.GoToSheet shtNeoGuiAcuut, "A1"
        Case "btnNAfspraken"
            ModSheets.GoToSheet shtNeoPrtAfspr, "A1"
        Case "btnNMedicatie"
            ModSheets.GoToSheet shtNeoPrtMedDisc, "A1"
        Case "btnNTPN"
            TPNAdviesNEO
        Case "btnNApotheek"
            ModSheets.GoToSheet shtNeoPrtApoth, "A1"
        Case "btnNWerkbrief"
            ModSheets.GoToSheet shtNeoPrtWerkbr, "A1"
        'grpDeveloper2
        Case "btnDeveloperMode"
            SetToDevelopmentMode
        Case "btnNaamGeven"
            ModMenuItems.GiveNameToRange
        Case "btnToggleLogging"
    End Select

    ' Waarom wordt dit aangeroepen na een menu item keuze???
    ' HideBars
    
    Application.ScreenUpdating = True
    
End Sub

Sub GetVisiblePed(control As IRibbonControl, ByRef blnVisible)

    Dim strPath As String

    strPath = LCase(Application.ActiveWorkbook.Path)
    
    If InStr(1, strPath, LCase(CONST_PELI_FOLDERNAME)) > 0 Or InStr(1, strPath, LCase(CONST_DEVELOP_FOLDERNAME)) > 0 Then
        blnVisible = True
    Else
        blnVisible = False
    End If
    
End Sub

Sub GetVisibleNeo(control As IRibbonControl, ByRef blnVisible)
    
    Dim strPath As String

    strPath = LCase(Application.ActiveWorkbook.Path)
    
    If InStr(1, strPath, LCase(CONST_NEO_FOLDERNAME)) > 0 Or InStr(1, strPath, LCase(CONST_DEVELOP_FOLDERNAME)) > 0 Then
        blnVisible = True
    Else
        blnVisible = False
    End If
    
End Sub

Sub GetVisibleDeveloper(control As IRibbonControl, ByRef blnVisible)

    Dim strPath As String

    strPath = LCase(Application.ActiveWorkbook.Path)
    
    If InStr(1, strPath, LCase(CONST_DEVELOP_FOLDERNAME)) > 0 Then
        blnVisible = True
    Else
        blnVisible = False
    End If
    
End Sub

