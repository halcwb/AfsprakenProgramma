Attribute VB_Name = "modPedContIV"
Option Explicit

Sub SetMedContIVToStandardItem(intRegel As Integer)

    Dim strMedicament As String
    Dim strMedSterkte As String
    Dim strOplHoev As String
    Dim strOplossing As String
    Dim varOplossing As Variant
    Dim strStand As String
    
    On Error GoTo SetMedContIVToStandardError

    strMedicament = "MedIVKeuze_" & intRegel
    strMedSterkte = "MedIVSterkte_" & intRegel
    strOplHoev = "MedIVMlOpl_" & intRegel
    strOplossing = "MedIVOplVlst_" & intRegel
    strStand = "MedIVStand_" & intRegel
    
    ModRange.SetRangeValue strMedSterkte, 0
    ModRange.SetRangeValue strOplHoev, 0
    ModRange.SetRangeValue strStand, 0
    
    If ModRange.GetRangeValue(strMedicament, 0) = 1 Then
        ModRange.SetRangeValue strOplossing, 1
    Else
        varOplossing = Application.VLookup(Range("tblMedicationContIV").Cells(Range(strMedicament).Value, 1), Range("tblMedicationContIV"), 22, False)
        ModRange.SetRangeValue strOplossing, varOplossing
    End If
    
    Exit Sub
    
SetMedContIVToStandardError:

    ModLog.LogError "SetMedContIVToStandard: " & " Error for intRegel " & intRegel

End Sub

Public Sub SetMedContIVToStand_01()
    
    SetMedContIVToStandardItem 1

End Sub

Public Sub SetMedContIVToStand_02()
    
    SetMedContIVToStandardItem 2

End Sub

Public Sub SetMedContIVToStand_03()
    
    SetMedContIVToStandardItem 3

End Sub

Public Sub SetMedContIVToStand_04()
        
    SetMedContIVToStandardItem 4

End Sub

Public Sub SetMedContIVToStand_05()
    
    SetMedContIVToStandardItem 5

End Sub

Public Sub SetMedContIVToStand_06()
    
    SetMedContIVToStandardItem 6

End Sub

Public Sub SetMedContIVToStand_07()

    SetMedContIVToStandardItem 7

End Sub

Public Sub SetMedContIVToStand_08()

    SetMedContIVToStandardItem 8

End Sub

Public Sub SetMedContIVToStand_09()
    
    SetMedContIVToStandardItem 9

End Sub

Public Sub SetMedContIVToStand_10()
    
    SetMedContIVToStandardItem 10

End Sub

Private Sub OpenInvoerNumeriek(intRegel As Integer, strRange As String, strUnit As String, intColumn As Integer)

    Dim frmInvoer As New FormInvoerNumeriek
    Dim varKeuze As Variant
    
    On Error GoTo OpenInvoerNumeriekError
    
    varKeuze = ModRange.GetRangeValue("MedIVKeuze_" & intRegel, vbNullString)
    
    With frmInvoer
        .Caption = "Medicament " & intRegel
        .lblParameter = "Oplossing"
        .lblEenheid = strUnit
        If ModRange.GetRangeValue("MedIVMlOpl_" & intRegel, 0) = 0 Then
            .txtWaarde = Application.WorksheetFunction.Index(Range("tblMedicationContIV"), varKeuze, 12)
        Else
            .txtWaarde = ModRange.GetRangeValue(strRange & intRegel, vbNullString)
        End If
        .Show
        If IsNumeric(.txtWaarde) Then
            If CDbl(.txtWaarde) = Application.WorksheetFunction.Index(Range("tblMedicationContIV"), varKeuze, 12) Then
                ModRange.SetRangeValue strRange & intRegel, 0
            Else
                ModRange.SetRangeValue strRange & intRegel, .txtWaarde
            End If
        End If
    End With
    
    Set frmInvoer = Nothing
    
    Exit Sub
    
OpenInvoerNumeriekError:

    ModLog.LogError "OpenInvoerNumeriek(" & Join(Array(intRegel, strRange, strUnit, intColumn), ", ") & ")"
    Set frmInvoer = Nothing

End Sub

Private Sub SetMedContIVSterkteItem(intRegel As Integer)

    Dim strUnit As String
    
    On Error GoTo SetMedContIVSterkteError

    strUnit = Application.WorksheetFunction.Index(Range("tblMedicationContIV"), Range("MedIVKeuze_" & intRegel), 4)
    OpenInvoerNumeriek intRegel, "MedIVSterkte_", strUnit, 11
    
    Exit Sub
    
SetMedContIVSterkteError:

    ModLog.LogError "SetMedContIVSterkteItem(" & intRegel & ")"

End Sub

Public Sub SetMedContIVSterkte_01()
    
    SetMedContIVSterkteItem 1

End Sub

Public Sub SetMedContIVSterkte_02()
    
    SetMedContIVSterkteItem 2

End Sub

Public Sub SetMedContIVSterkte_03()
    
    SetMedContIVSterkteItem 3

End Sub

Public Sub SetMedContIVSterkte_04()
    
    SetMedContIVSterkteItem 4

End Sub

Public Sub SetMedContIVSterkte_05()
    
    SetMedContIVSterkteItem 5

End Sub

Public Sub SetMedContIVSterkte_06()
    
    SetMedContIVSterkteItem 6

End Sub

Public Sub SetMedContIVSterkte_07()
    
    SetMedContIVSterkteItem 7

End Sub

Public Sub SetMedContIVSterkte_08()
    
    SetMedContIVSterkteItem 8

End Sub

Public Sub SetMedContIVSterkte_09()
    
    SetMedContIVSterkteItem 9

End Sub

Public Sub SetMedContIVSterkte_10()
    
    SetMedContIVSterkteItem 10

End Sub

Private Sub SetMedContIVOplossingItem(intRegel As Integer)

    OpenInvoerNumeriek intRegel, "MedIVMlOpl_", "mL", 12

End Sub

Public Sub SetMedContIVOplossing_01()
    
    SetMedContIVOplossingItem 1

End Sub

Public Sub SetMedContIVOplossing_02()
    
    SetMedContIVOplossingItem 2

End Sub

Public Sub SetMedContIVOplossing_03()
    
    SetMedContIVOplossingItem 3

End Sub

Public Sub SetMedContIVOplossing_04()
    
    SetMedContIVOplossingItem 4

End Sub

Public Sub SetMedContIVOplossing_05()
    
    SetMedContIVOplossingItem 5

End Sub

Public Sub SetMedContIVOplossing_06()
    
    SetMedContIVOplossingItem 6

End Sub

Public Sub SetMedContIVOplossing_07()
    
    SetMedContIVOplossingItem 7

End Sub

Public Sub SetMedContIVOplossing_08()
    
    SetMedContIVOplossingItem 8

End Sub

Public Sub SetMedContIVOplossing_09()
    
    SetMedContIVOplossingItem 9

End Sub

Public Sub SetMedContIVOplossing_10()
    
    SetMedContIVOplossingItem 10

End Sub

