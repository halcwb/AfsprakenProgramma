Attribute VB_Name = "ModFormularium"
Option Explicit

Private m_Formularium As ClassFormularium
Private m_FormConfig As ClassFormConfig

Private Const constFormularium As String = "FormulariumDb.xlsm"
Private Const constFormDbStart As Integer = 3

'--- Formularium ---
Private Const constGPKIndx As Integer = 1
Private Const constATCIndx As Integer = 2
Private Const constHoofdGroepIndx As Integer = 3
Private Const constSubGroepIndx As Integer = 4
Private Const constGeneriekIndx As Integer = 5
Private Const constProductIndx As Integer = 6
Private Const constEtiketIndx As Integer = 7
Private Const constVormIndx As Integer = 8
Private Const constRouteIndx As Integer = 9
Private Const constSterkteIndx As Integer = 10
Private Const constEenheidIndx As Integer = 11
Private Const constStandDoseIndx As Integer = 12
Private Const constDoseEenheidIndx As Integer = 13
Private Const constIndicatiesIndx As Integer = 14
Private Const constFreqIndx As Integer = 15
Private Const constPICU_DoseIndx As Integer = 16
Private Const constPICU_OnderIndx As Integer = 17
Private Const constPICU_BovenIndx As Integer = 18
Private Const constNICU_DoseIndx As Integer = 19
Private Const constNICU_OnderIndx As Integer = 20
Private Const constNICU_BovenIndx As Integer = 21
Private Const constMaxDoseIndx As Integer = 22
Private Const constPICU_MaxConcIndx As Integer = 23
Private Const constPICU_OplVlstIndx As Integer = 24
Private Const constPICU_OplVolIndx As Integer = 25
Private Const constPICU_MinTijdIndx As Integer = 26
Private Const constNICU_MaxConcIndx As Integer = 27
Private Const constNICU_OplVlstIndx As Integer = 28
Private Const constNICU_OplVolIndx As Integer = 29
Private Const constNICU_MinTijdIndx As Integer = 30

Public Sub Formularium_Initialize()

    Dim intN As Integer
    Dim intC As Integer
    Dim strTitle As String

    If Not m_Formularium Is Nothing Then Exit Sub

    strTitle = "Formularium wordt geladen, een ogenblik geduld a.u.b. ..."
    
    ModProgress.StartProgress strTitle
       
    Set m_Formularium = New ClassFormularium
    m_Formularium.GetMedicamenten True
    
    ModProgress.FinishProgress

End Sub

Private Sub FormConfig_Initialize()

    Dim intN As Integer
    Dim intC As Integer
    Dim strTitle As String

    If Not m_FormConfig Is Nothing Then Exit Sub

    strTitle = "Formularium configuratie wordt geladen, een ogenblik geduld a.u.b. ..."
    
    ModProgress.StartProgress strTitle
       
    Set m_FormConfig = New ClassFormConfig
    m_FormConfig.GetMedicamenten True
      
    ModProgress.FinishProgress

End Sub

Public Function Formularium_IsInitialized() As Boolean

    Formularium_IsInitialized = Not m_Formularium Is Nothing

End Function

Public Function Formularium_GetFormularium() As ClassFormularium

    Formularium_Initialize
    Set Formularium_GetFormularium = m_Formularium

End Function

Public Function Formularium_GetFormConfig() As ClassFormConfig

    FormConfig_Initialize
    Set Formularium_GetFormConfig = m_FormConfig

End Function

Private Sub Test_Formularium_Initialize()

    Formularium_Initialize

End Sub

Public Sub Formularium_GetMedicamenten(objFormularium As ClassFormularium, ByVal blnShowProgress As Boolean)

    Dim intN As Integer
    Dim intC As Integer
    Dim objFormRange As Range
    Dim objSheet As Worksheet
    
    Dim strFileName As String
    Dim strName As String
    Dim strSheet As String
    Dim blnIsPed As Boolean
    
    Dim objMed As ClassMedicatieDisc
    
    On Error GoTo GetMedicamentenError
    
    blnIsPed = MetaVision_IsPediatrie()
    
    strName = constFormularium
    strSheet = "Table"
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    strFileName = ModMedDisc.GetFormulariumDatabasePath() + strName

    Workbooks.Open strFileName, True, True
    
    Set objSheet = Workbooks(strName).Worksheets(strSheet)
    Set objFormRange = objSheet.Range("A1").CurrentRegion
        
    intC = objFormRange.Rows.Count
    For intN = constFormDbStart To intC
        Set objMed = New ClassMedicatieDisc
        
        With objMed
            
            .GPK = objFormRange.Cells(intN, constGPKIndx).Value2
            .MainGroup = objFormRange.Cells(intN, constHoofdGroepIndx).Value2
            .SubGroup = objFormRange.Cells(intN, constSubGroepIndx).Value2
            
            .ATC = objFormRange.Cells(intN, constATCIndx).Value2
            .Generic = objFormRange.Cells(intN, constGeneriekIndx).Value2
            .Product = objFormRange.Cells(intN, constProductIndx).Value2
            .Shape = objFormRange.Cells(intN, constVormIndx).Value2
            .GenericQuantity = objFormRange.Cells(intN, constSterkteIndx).Value2
            .GenericUnit = objFormRange.Cells(intN, constEenheidIndx).Value2
            .Label = objFormRange.Cells(intN, constEtiketIndx).Value2
            .MultipleQuantity = objFormRange.Cells(intN, constStandDoseIndx).Value2
            .MultipleUnit = objFormRange.Cells(intN, constDoseEenheidIndx).Value2
            
            .SetRouteList objFormRange.Cells(intN, constRouteIndx).Value2
            .SetIndicationList objFormRange.Cells(intN, constIndicatiesIndx).Value2
            .SetFreqList objFormRange.Cells(intN, constFreqIndx).Value2
            
            .NormDose = objFormRange.Cells(intN, IIf(blnIsPed, constPICU_DoseIndx, constNICU_DoseIndx)).Value2
            .MinDose = objFormRange.Cells(intN, IIf(blnIsPed, constPICU_OnderIndx, constNICU_OnderIndx)).Value2
            .MaxDose = objFormRange.Cells(intN, IIf(blnIsPed, constPICU_BovenIndx, constNICU_BovenIndx)).Value2
            .AbsMaxDose = objFormRange.Cells(intN, constMaxDoseIndx).Value2
            
            .MaxConc = objFormRange.Cells(intN, IIf(blnIsPed, constPICU_MaxConcIndx, constNICU_MaxConcIndx)).Value2
            .Solution = objFormRange.Cells(intN, IIf(blnIsPed, constPICU_OplVlstIndx, constNICU_OplVlstIndx)).Value2
            .SolutionVolume = objFormRange.Cells(intN, IIf(blnIsPed, constPICU_OplVolIndx, constNICU_OplVolIndx)).Value2
            .MinInfusionTime = objFormRange.Cells(intN, IIf(blnIsPed, constPICU_MinTijdIndx, constNICU_MinTijdIndx)).Value2
        
        End With
                
        objFormularium.AddMedicament objMed
        
        If blnShowProgress Then ModProgress.SetJobPercentage "Formularium laden", intC, intN
        
    Next intN
    
    Workbooks(strName).Close

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Exit Sub
    
GetMedicamentenError:
    
    ModLog.LogError Err, "Could not retrieve medicament from: " & strFileName
    
    On Error Resume Next
    
    Workbooks(strName).Close

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    ModProgress.FinishProgress
    
End Sub

Public Sub Formularium_GetMedDiscConfig(objFormularium As ClassFormConfig, ByVal blnShowProgress As Boolean)

    Dim intN As Integer
    Dim intC As Integer
    Dim objFormRange As Range
    Dim objSheet As Worksheet
    
    Dim strFileName As String
    Dim strName As String
    Dim strSheet As String
    Dim blnIsPed As Boolean
    
    Dim objMed As ClassMedDiscConfig
    
    On Error GoTo GetMedicamentenError
    
    blnIsPed = MetaVision_IsPediatrie()
    
    strName = constFormularium
    strSheet = "Table"
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    strFileName = ModMedDisc.GetFormulariumDatabasePath() + strName

    Workbooks.Open strFileName, True, True
    
    Set objSheet = Workbooks(strName).Worksheets(strSheet)
    Set objFormRange = objSheet.Range("A1").CurrentRegion
        
    intC = objFormRange.Rows.Count
    For intN = constFormDbStart To intC
        Set objMed = New ClassMedDiscConfig
        
        With objMed
            .GPK = objFormRange.Cells(intN, constGPKIndx).Value2
            .MainGroup = objFormRange.Cells(intN, constHoofdGroepIndx).Value2
            .SubGroup = objFormRange.Cells(intN, constSubGroepIndx).Value2
            
            .ATC = objFormRange.Cells(intN, constATCIndx).Value2
            .Generic = objFormRange.Cells(intN, constGeneriekIndx).Value2
            .Product = objFormRange.Cells(intN, constProductIndx).Value2
            .Shape = objFormRange.Cells(intN, constVormIndx).Value2
            .GenericQuantity = objFormRange.Cells(intN, constSterkteIndx).Value2
            .GenericUnit = objFormRange.Cells(intN, constEenheidIndx).Value2
            .Label = objFormRange.Cells(intN, constEtiketIndx).Value2
            .MultipleQuantity = objFormRange.Cells(intN, constStandDoseIndx).Value2
            .MultipleUnit = objFormRange.Cells(intN, constDoseEenheidIndx).Value2
            
            .SetRouteList objFormRange.Cells(intN, constRouteIndx).Value2
            .SetIndicationList objFormRange.Cells(intN, constIndicatiesIndx).Value2
            .SetFreqList objFormRange.Cells(intN, constFreqIndx).Value2
            
            .NeoNormDose = objFormRange.Cells(intN, constNICU_DoseIndx).Value2
            .NeoMinDose = objFormRange.Cells(intN, constNICU_OnderIndx).Value2
            .NeoMaxDose = objFormRange.Cells(intN, constNICU_BovenIndx).Value2
            
            .PedNormDose = objFormRange.Cells(intN, constPICU_DoseIndx).Value2
            .PedMinDose = objFormRange.Cells(intN, constPICU_OnderIndx).Value2
            .PedMaxDose = objFormRange.Cells(intN, constPICU_BovenIndx).Value2
            
            .PedAbsMaxDose = objFormRange.Cells(intN, constMaxDoseIndx).Value2
            
            .PedMaxConc = objFormRange.Cells(intN, constPICU_MaxConcIndx).Value2
            .PedSolutionVolume = objFormRange.Cells(intN, constPICU_OplVolIndx).Value2
            .PedSolution = objFormRange.Cells(intN, constPICU_OplVlstIndx).Value2
            .PedMinInfusionTime = objFormRange.Cells(intN, constPICU_MinTijdIndx).Value2
            
            .NeoMaxConc = objFormRange.Cells(intN, constNICU_MaxConcIndx).Value2
            .NeoSoutionVolume = objFormRange.Cells(intN, constNICU_OplVolIndx).Value2
            .NeoSolution = objFormRange.Cells(intN, constNICU_OplVlstIndx).Value2
            .NeoMinInfustionTime = objFormRange.Cells(intN, constNICU_MinTijdIndx).Value2
            
        End With
        
        objFormularium.AddMedicament objMed
        
        If blnShowProgress Then ModProgress.SetJobPercentage "Formularium laden", intC, intN
        
    Next intN
    
    Workbooks(strName).Close

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Exit Sub
    
GetMedicamentenError:
    
    ModLog.LogError Err, "Could not retrieve medicament from: " & strFileName
    
    On Error Resume Next
    
    Workbooks(strName).Close

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    ModProgress.FinishProgress
    
End Sub

Public Sub Formularium_SaveMedDiscConfig(ByVal blnShowProgress As Boolean)

    Dim intN As Integer
    Dim intC As Integer
    Dim objFormRange As Range
    Dim objSheet As Worksheet
    
    Dim strFileName As String
    Dim strName As String
    Dim strSheet As String
    Dim blnIsPed As Boolean
    
    Dim objMed As ClassMedDiscConfig
    
    On Error GoTo GetMedicamentenError
    
    blnIsPed = MetaVision_IsPediatrie()
    
    strName = constFormularium
    strSheet = "Table"
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    strFileName = ModMedDisc.GetFormulariumDatabasePath() + strName

    Workbooks.Open strFileName, True, False
    
    Set objSheet = Workbooks(strName).Worksheets(strSheet)
    Set objFormRange = objSheet.Range("A1").CurrentRegion
        
    intC = objFormRange.Rows.Count
    For intN = constFormDbStart To intC
        Set objMed = m_FormConfig.GPK(objFormRange.Cells(intN, constGPKIndx).Value2)
        
        With objMed
            
            If Not Trim(LCase(objFormRange.Cells(intN, constGeneriekIndx).Value2)) Then
                objFormRange.Cells(intN, constGeneriekIndx).Value2 = objMed.Generic
            End If
            
            If .MultipleQuantity > 0 Then objFormRange.Cells(intN, constStandDoseIndx).Value2 = .MultipleQuantity
            If Not .MultipleUnit = vbNullString Then objFormRange.Cells(intN, constDoseEenheidIndx).Value2 = .MultipleUnit
            
            If Not .GetFreqListString() = vbNullString Then objFormRange.Cells(intN, constFreqIndx).Value2 = .GetFreqListString()
            
            If .NeoNormDose > 0 Then objFormRange.Cells(intN, constNICU_DoseIndx).Value2 = .NeoNormDose
            If .NeoMinDose > 0 Then objFormRange.Cells(intN, constNICU_OnderIndx).Value2 = .NeoMinDose
            If .NeoMaxDose > 0 Then objFormRange.Cells(intN, constNICU_BovenIndx).Value2 = .NeoMaxDose
            
            If .PedNormDose > 0 Then objFormRange.Cells(intN, constPICU_DoseIndx).Value2 = .PedNormDose
            If .PedMinDose > 0 Then objFormRange.Cells(intN, constPICU_OnderIndx).Value2 = .PedMinDose
            If .PedMaxDose > 0 Then objFormRange.Cells(intN, constPICU_BovenIndx).Value2 = .PedMaxDose
            
            If .PedAbsMaxDose > 0 Then objFormRange.Cells(intN, constMaxDoseIndx).Value2 = .PedAbsMaxDose
            
            If .PedMaxConc > 0 Then objFormRange.Cells(intN, constPICU_MaxConcIndx).Value2 = .PedMaxConc
            If .PedSolutionVolume > 0 Then objFormRange.Cells(intN, constPICU_OplVolIndx).Value2 = .PedSolutionVolume
            If Not .PedSolution = vbNullString Then objFormRange.Cells(intN, constPICU_OplVlstIndx).Value2 = .PedSolution
            If .PedMinInfusionTime > 0 Then objFormRange.Cells(intN, constPICU_MinTijdIndx).Value2 = .PedMinInfusionTime
        
            If .NeoMaxConc > 0 Then objFormRange.Cells(intN, constNICU_MaxConcIndx).Value2 = .NeoMaxConc
            If .NeoSoutionVolume > 0 Then objFormRange.Cells(intN, constNICU_OplVolIndx).Value2 = .NeoSoutionVolume
            If Not .NeoSolution = vbNullString Then objFormRange.Cells(intN, constNICU_OplVlstIndx).Value2 = .NeoSolution
            If .NeoMinInfustionTime > 0 Then objFormRange.Cells(intN, constNICU_MinTijdIndx).Value2 = .NeoMinInfustionTime
        
        End With
        
        m_FormConfig.AddMedicament objMed
        
        If blnShowProgress Then ModProgress.SetJobPercentage "Formularium laden", intC, intN
    Next intN
    
    Workbooks(strName).Save
    Workbooks(strName).Close
    
    Set m_Formularium = Nothing

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Exit Sub
    
GetMedicamentenError:
    
    ModLog.LogError Err, "Could not retrieve medicament from: " & strFileName
    
    On Error Resume Next
    
    Workbooks(strName).Close

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    ModProgress.FinishProgress
    
End Sub

Public Sub Formularium_ShowConfig()

    FormAdminMedDisc.Show
    
End Sub
