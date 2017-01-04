Attribute VB_Name = "ModRibbon"
Option Explicit

Public Sub ButtonOnAction(ctrlMenuItem As IRibbonControl)
'
' Code for onAction callback. Ribbon control button
'
    Application.ScreenUpdating = False
    
    Select Case ctrlMenuItem.Id
    
        'grpAfspraken                                       ' -- AFSPRAKEN --
        
        Case "btnClose"                                     ' -> Programma Afsluiten
            ModApplication.CloseAfspraken
        
        Case "btnClear"                                     ' -> Alles Verwijderen
            ModProgress.StartProgress "Patient Data Verwijderen"
            ModPatient.ClearPatient True, True
            ModProgress.FinishProgress
            ModSheet.SelectPedOrNeoStartSheet
        
        'grpBedden                                          ' -- BEDDEN --
        
        Case "btnOpenBed"                                   ' -> Bed Openen
            ModBed.OpenBed
            ModSheet.SelectPedOrNeoStartSheet
        
        Case "btnSaveBed"                                   ' -> Bed Opslaan
            ModBed.CloseBed (True)
            ModSheet.SelectPedOrNeoStartSheet
        
        Case "btnEnterPatient"                              ' -> Patient Gegevens
            ModPatient.EnterPatientDetails
            
        'grpPediatrie                                       ' -- PEDIATRIE --
        
        Case "btnPedMedIV"                                  ' -> Continue IV Medicatie
            ModSheet.GoToSheet shtPedGuiMedIV, "A6"
        
        Case "btnPedMedDisc"                                ' -> Discontinue Medicatie
            ModSheet.GoToSheet shtGlobGuiMedDisc, "A6"
        
        Case "btnPedIVandPM"                                ' -> Lijnen en PM
            ModSheet.GoToSheet shtPedGuiLijnPM, "A6"
        
        Case "btnPedEntTPN"                                 ' -> Voeding en TPN
            ModSheet.GoToSheet shtPedGuiEntTPN, "A6"
        
        Case "btnPedLab"                                    ' -> Lab Aanvragen
            ModSheet.GoToSheet shtPedGuiLab, "A1"
        
        Case "btnPedExtra"                                  ' -> Afspraken en Controles
            ModSheet.GoToSheet shtPedGuiAfspr, "A6"
            
        'grpNeonatologie                                    ' -- NEONATOLOGIE --
        
        Case "btnNeoMedIV"                                  ' -> Infuusbrief
            ModSheet.GoToSheet shtNeoGuiInfB, "A9"
        
        Case "btnNeoMedDisc"                                ' -> Discontinue Medicatie
            ModSheet.GoToSheet shtGlobGuiMedDisc, "A6"
        
        Case "btnNeoExtra"                                  ' -> Afspraken en Controles
            ModSheet.GoToSheet shtNeoGuiAfspr, "A6"
        
        Case "btnNTPNadvies"
            ModNeoInfB.NeoInfB_TPNAdvice
        
        Case "btnNeoLab"                                    ' -> Lab Aanvragen
            ModSheet.GoToSheet shtNeoGuiLab, "A1"
        
        Case "btnNeo1700"                                   ' -> Infuusbrief 17:00
            ModSheet.GoToSheet shtNeoGui1700, "A9"
        
        'grpRemoveData                                      ' -- ACTIES --
                
        Case "btnRemoveLab"                                 ' Lab Verwijderen
            ModClear.ClearLab
            ModSheet.SelectPedOrNeoLabSheet
        
        Case "btnRemoveExtra"                               ' Afspraken Controles Verwijderen
            ModClear.ClearAfspraken
            ModSheet.SelectPedOrNeoAfsprSheet
        
        ' grpNeoActions                                     ' -- INFUUSBRIEF OVERZETTEN --
        
        Case "btnCopy1700"                                  ' -> 17:00 uur Overzetten
            ModNeoInfB.CopyToActueel
        
        Case "btnCopyCurrent"                               ' -> Actueel Overzetten
            ModNeoInfB.NeoInfB_CopyAfspraken
            
        'grpPedPrint                                        ' -- PRINT PEDIATRIE --
        
        Case "btnPedPrintAcuut"                             ' -> Acute Blad
            ModSheet.GoToSheet shtPedGuiAcuut, "A1"
        
        Case "btnPedPrintMedIV"                             ' -> Medicatie IV
            ModSheet.GoToSheet shtPedPrtAfspr, "A1"
        
        Case "btnPedPrintMedDisc"                           ' -> Medicatie Discontinu
            ModSheet.GoToSheet shtPedPrtMedDisc, "A1"
        
        Case "btnPedPrintTPN"                               ' -> TPN Brief
            ModPedEntTPN.PedEntTPN_SelectTPNPrint
            
        'grpNeoPrint                                        ' -- PRINT NEO ---
        
        Case "btnNeoPrintAcuut"                             ' -> Acute Blad
            ModSheet.GoToSheet shtNeoGuiAcuut, "A1"
        
        Case "btnNeoPrintMedIV"                             ' -> Infuus Brief
            ModSheet.GoToSheet shtNeoPrtAfspr, "A1"
        
        Case "btnNeoPrintMedDisc"                           ' -> Medicatie Discontinu
            ModSheet.GoToSheet shtNeoPrtMedDisc, "A1"
        
        Case "btnNTPN"
            ModNeoInfB.NeoInfB_TPNAdvice
        
        Case "btnNeoPrintApoth"                             ' -> Apotheek
            ModSheet.GoToSheet shtNeoPrtApoth, "A1"
        
        Case "btnNeoPrintWerkbr"                            ' -> Werkbrief
            ModSheet.GoToSheet shtNeoPrtWerkbr, "A1"
            
        'grpDevelopment                                     ' -- DEVELOPMENT --
        
        Case "btnDevMode"                                   ' -> Development Mode
            ModApplication.SetToDevelopmentMode
        
        Case "btnToggleLogging"                             ' -> Toggle Logging
            ModSetting.ToggleLogging
    
        Case "btnRangeNames"                                ' -> Range Names
            ModRange.GiveNameToRange
        
        Case Else
            MsgBox ctrlMenuItem.Id & " has no select case", vbCritical
            
        
    End Select

    Application.ScreenUpdating = True
    
End Sub

Public Sub GetVisiblePed(ctrContr As IRibbonControl, ByRef blnVisible)

    Dim strPath, strPedDir As String
    Dim blnIsDevelop

    blnIsDevelop = ModSetting.IsDevelopmentMode()
    strPath = Application.ActiveWorkbook.Path
    strPedDir = ModSetting.GetPedDir()
    
    If ModString.ContainsCaseInsensitive(strPath, strPedDir) Or blnIsDevelop Then
        blnVisible = True
    Else
        blnVisible = False
    End If
    
End Sub

Public Sub GetVisibleNeo(ctrContr As IRibbonControl, ByRef blnVisible)
    
    Dim strPath, strPedDir As String
    Dim blnIsDevelop

    blnIsDevelop = ModSetting.IsDevelopmentMode()
    strPath = Application.ActiveWorkbook.Path
    strPedDir = ModSetting.GetNeoDir()
    
    If ModString.ContainsCaseInsensitive(strPath, strPedDir) Or blnIsDevelop Then
        blnVisible = True
    Else
        blnVisible = False
    End If
    
End Sub

Public Sub GetVisibleDevelopment(ctrContr As IRibbonControl, ByRef blnVisible)

    blnVisible = ModSetting.IsDevelopmentMode()
    
End Sub

