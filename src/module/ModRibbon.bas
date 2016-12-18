Attribute VB_Name = "ModRibbon"
Option Explicit

Public Sub ButtonOnAction(control As IRibbonControl)
'
' Code for onAction callback. Ribbon control button
'

    Application.ScreenUpdating = False
    
    Select Case control.ID
        'GroupAfspraken
        Case "btnAfsluiten"
            Afsluiten
        Case "btnAfsprakenVerwijderen"
            clearPat True
            OpenStartSheet
        'GroupBedden
        Case "btnBedOpenen"
            PaPatientenLijst
        Case "btnBedOpslaan"
            BeSluitBed
            OpenStartSheet
        Case "btnGegevensInvoeren"
            NieuwePatient
        'grpPediatrie
        Case "btnPContinueivmedicatie"
            gaNaarMedicatieIV
            Range("A6").Select
            ActiveWindow.ScrollRow = 1
        Case "btnPDiscontinuemedicatie"
            gaNaarMedicatieOverig
            Range("A6").Select
            ActiveWindow.ScrollRow = 1
        Case "btnPInfusen"
            gaNaarInfusen
            Range("A6").Select
            ActiveWindow.ScrollRow = 1
        Case "btnPIntake"
            gaNaarIntake
            Range("A6").Select
            ActiveWindow.ScrollRow = 1
        Case "btnPLaboratoriumbepalingen"
            gaNaarLab
            Range("A1").Select
            ActiveWindow.ScrollRow = 1
        Case "btnPAanvullendeAfspraken"
            gaNaarAanvullendeAfspraken
            Range("A6").Select
            ActiveWindow.ScrollRow = 1
        'grpNeonatologie
        Case "btnNInfuusbrief"
            GaNaarAfspraken
            Range("A9").Select
            ActiveWindow.ScrollRow = 1
        Case "btnNDiscontinuemedicatie"
            gaNaarMedicatieOverig
            Range("A6").Select
            ActiveWindow.ScrollRow = 1
        Case "btnNAanvullendeAfspraken"
            gaNaarAanvullendeAfsprakenNeo
            Range("A6").Select
            ActiveWindow.ScrollRow = 1
        Case "btnNTPNadvies"
            TPNAdviesNEO
        Case "btnNLaboratoriumbepalingen"
            gaNaarLabNeo
            Range("A1").Select
            ActiveWindow.ScrollRow = 1
        Case "btnNAfspraken1700"
            gaNaarAfspraken1700Neo
            Range("A9").Select
            ActiveWindow.ScrollRow = 1
        'grpGegevensVerwijderen -> Acties
        Case "btnNAfspraken1700Overzetten"
            ModAfspraken1700.CopyToActueel
            Range("A1").Select
            ActiveWindow.ScrollRow = 1
        Case "btnNActueleAfsprakenOverzetten"
            ModAfspraken.AfsprakenOvernemen
            Range("A1").Select
            ActiveWindow.ScrollRow = 1
        Case "btnLabVerwijderen"
            VerwijderLab
            OpenLabSheet
        Case "btnAanvullendVerwijderen"
            VerwijderAanvullendeAfspraken
            OpenAanvullendeSheet
        'grpPrintPediatrie
        Case "btnPAcuteBlad"
            gaNaarAcuteOpvang
            Range("A1").Select
            ActiveWindow.ScrollRow = 1
        Case "btnPPrintAfspraken"
            gaNaarAfspraakBlad
            Range("A1").Select
            ActiveWindow.ScrollRow = 1
        Case "btnPMedicatie"
            gaNaarMedicatie
            Range("A1").Select
            ActiveWindow.ScrollRow = 1
        Case "btnPTPN"
            gaNaarTPNblad
        'grpPrintNeonatologie
        Case "btnNAcuteBlad"
            gaNaarAcuteOpvangNeo
            Range("A1").Select
            ActiveWindow.ScrollRow = 1
        Case "btnNAfspraken"
            GaNaarPrint
            Range("A1").Select
            ActiveWindow.ScrollRow = 1
        Case "btnNMedicatie"
            gaNaarMedicatieNeo
            Range("A1").Select
            ActiveWindow.ScrollRow = 1
        Case "btnNTPN"
            TPNAdviesNEO
        Case "btnNApotheek"
            GaNaarApotheek
            Range("A1").Select
            ActiveWindow.ScrollRow = 1
        Case "btnNWerkbrief"
            GaNaarWerkBrief
            Range("A1").Select
            ActiveWindow.ScrollRow = 1
        'grpDeveloper2
        Case "btnDeveloperMode"
            SetToDevelopmentMode
        Case "btnToggleLogging"
    End Select

    HideBars
    
    Application.ScreenUpdating = True
End Sub

Sub GetVisiblePed(control As IRibbonControl, ByRef visible)
Dim strPath As String

    strPath = LCase(Application.ActiveWorkbook.Path)
    
    If InStr(1, strPath, LCase(CONST_PELI_FOLDERNAME)) > 0 Or InStr(1, strPath, LCase(CONST_DEVELOP_FOLDERNAME)) > 0 Then
        visible = True
    Else
        visible = False
    End If
End Sub

Sub GetVisibleNeo(control As IRibbonControl, ByRef visible)
Dim strPath As String

    strPath = LCase(Application.ActiveWorkbook.Path)
    
    If InStr(1, strPath, LCase(CONST_NEO_FOLDERNAME)) > 0 Or InStr(1, strPath, LCase(CONST_DEVELOP_FOLDERNAME)) > 0 Then
        visible = True
    Else
        visible = False
    End If
End Sub

Sub GetVisibleDeveloper(control As IRibbonControl, ByRef visible)
Dim strPath As String

    strPath = LCase(Application.ActiveWorkbook.Path)
    
    If InStr(1, strPath, LCase(CONST_DEVELOP_FOLDERNAME)) > 0 Then
        visible = True
    Else
        visible = False
    End If
End Sub

