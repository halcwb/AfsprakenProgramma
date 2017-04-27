Attribute VB_Name = "ModRibbon"
Option Explicit

' Module to handle ribbon events
' Note: blnVisible in GetVisible functions is not a boolean!! Using blnVisible As Boolean will result in a type mismatch!

Public Sub ButtonOnAction(ByRef ctrlMenuItem As IRibbonControl)
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
            ModPatient.PatientClearAll True, True
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
            ModNeoInfB.NeoInfB_SelectInfB False
        
        Case "btnNeoMedDisc"                                ' -> Discontinue Medicatie
            ModSheet.GoToSheet shtGlobGuiMedDisc, "A6"
        
        Case "btnNeoExtra"                                  ' -> Afspraken en Controles
            ModSheet.GoToSheet shtNeoGuiAfspr, "A6"
        
        Case "btnNTPNadvies"
            ModNeoInfB.NeoInfB_TPNAdvice
        
        Case "btnNeoLab"                                    ' -> Lab Aanvragen
            ModSheet.GoToSheet shtNeoGuiLab, "A1"
        
        Case "btnNeo1700"                                   ' -> Infuusbrief 17:00
            ModNeoInfB.NeoInfB_SelectInfB True
        
        'grpRemoveData                                      ' -- ACTIES --
                
        Case "btnRemoveLab"                                 ' Lab Verwijderen
            ClearLab
            ModSheet.SelectPedOrNeoLabSheet
        
        Case "btnRemoveExtra"                               ' Afspraken Controles Verwijderen
            ClearAfspraken
            ModSheet.SelectPedOrNeoAfsprSheet
        
        ' grpNeoActions                                     ' -- INFUUSBRIEF OVERZETTEN --
        
        Case "btnCopy1700"                                  ' -> 17:00 uur Overzetten
            ModNeoInfB.NeoInfB_ShowFormCopy1700ToAct
        
        Case "btnCopyCurrent"                               ' -> Actueel Overzetten
            ModNeoInfB.NeoInfB_CopyActTo1700
            
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
              
        Case "btnNeoPrintApoth"                             ' -> Apotheek
            ' ModSheet.GoToSheet shtNeoPrtApoth, "A1"
            ModNeoPrint.PrintApotheekWerkBrief
        
        Case "btnNeoPrintWerkbr"                            ' -> Werkbrief
            ' ModSheet.GoToSheet shtNeoPrtWerkbr, "A1"
            
        'grpDevelopment                                     ' -- DEVELOPMENT --
        
        Case "btnDevMode"                                   ' -> Development Mode
            ModApplication.SetToDevelopmentMode
        
        Case "btnToggleLogging"                             ' -> Toggle Logging
            ModSetting.ToggleLogging
    
        Case "btnRangeNames"                                ' -> Name Range
            ModRange.GiveNameToRange
            
        Case "btnWriteNames"                                ' -> Write Names
            ModRange.WriteNamesToGlobNames
            
        Case "btnReplaceNames"                              ' -> Replace Names
            ModRange.ReplaceRangeNames
            
        Case "btnRefreshPatientData"                        ' -> Refresh Patient Data
            ModRange.RefreshPatientData
            
        Case "btnExportSource"                              ' -> Export To Source
            ModUtils.ExportForSourceControl
        
        'grpAdmin                                           ' -- ADMISTRATION --
        
        Case "btnOpenSettings"                              ' -> Instellingen
             ModMessage.ShowMsgBoxExclam "Nog niet geimplementeerd"
        
        Case "btnSetColors"                                 ' -> Kleuren Instellen
             ModAdmin.ShowColorPicker
        
        Case "btnCreatePedData"                             ' -> Pediatrie DataFiles
             ModAdmin.SetUpPedDataDir
        
        Case "btnCreateNeoData"                             ' -> Neonatologie DataFiles
             ModAdmin.SetUpNeoDataDir
        
        Case Else
            ModMessage.ShowMsgBoxError ctrlMenuItem.Id & " has no select case"
            
        
    End Select

    Application.ScreenUpdating = True
    
End Sub

Public Sub RibbonOnLoad(ByRef objRibbon As IRibbonUI)

    ModLog.LogInfo "RibbonOnLoad"

End Sub

Public Sub GetVisiblePed(ByRef ctrContr As IRibbonControl, ByRef blnVisible As Variant)

    Dim strPath As String
    Dim strPedDir As String
    Dim blnIsDevelop As Boolean

    blnIsDevelop = ModSetting.IsDevelopmentDir()
    strPath = Application.ActiveWorkbook.Path
    strPedDir = ModSetting.GetPedDir()
    
    If ModString.ContainsCaseInsensitive(strPath, strPedDir) Or blnIsDevelop Then
        blnVisible = True
    Else
        blnVisible = False
    End If
    
End Sub

Public Sub GetVisibleNeo(ByRef ctrContr As IRibbonControl, ByRef blnVisible As Variant)
    
    Dim strPath As String
    Dim strPedDir As String
    Dim blnIsDevelop As Boolean

    blnIsDevelop = ModSetting.IsDevelopmentDir()
    strPath = Application.ActiveWorkbook.Path
    strPedDir = ModSetting.GetNeoDir()
    
    If ModString.ContainsCaseInsensitive(strPath, strPedDir) Or blnIsDevelop Then
        blnVisible = True
    Else
        blnVisible = False
    End If
    
End Sub

Public Sub GetVisibleDevelopment(ByRef ctrContr As IRibbonControl, ByRef blnVisible As Variant)

    blnVisible = ModSetting.IsDevelopmentDir()
    
End Sub

Public Sub GetVisibleAdmin(ByRef ctrContr As IRibbonControl, ByRef blnVisible As Variant)

    blnVisible = True ' ModSetting.IsDevelopmentMode()
    
End Sub

Private Sub ClearLab()
    
    If ModSetting.IsDevelopmentDir Then
        ModNeoLab.NeoLab_Clear
        ModPedLab.PedLab_Clear
    Else
        If ModApplication.IsNeoDir() Then ModNeoLab.NeoLab_Clear
        If ModApplication.IsPedDir() Then ModPedLab.PedLab_Clear
    End If
    
End Sub

Private Sub ClearAfspraken()

    If ModSetting.IsDevelopmentDir Then
        ModNeoAfspr.NeoAfspr_Clear
        ModPedAfspr.PedAfspr_Clear
    Else
        If ModApplication.IsNeoDir() Then ModNeoAfspr.NeoAfspr_Clear
        If ModApplication.IsPedDir() Then ModPedAfspr.PedAfspr_Clear
    End If

End Sub


