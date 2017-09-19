Attribute VB_Name = "ModFormularium"
Option Explicit

Private m_Formularium As ClassFormularium

Private Const constFormularium As String = "FormulariumDb.xlsm"
Private Const constFormDbStart As Integer = 3

'--- Formularium ---
Private Const constGPKIndx As Integer = 1
Private Const constATCIndx As Integer = 2
Private Const constHoofdGroepIndx As Integer = 3
Private Const constSubGroepIndx As Integer = 4
Private Const constGeneriekIndx As Integer = 5
Private Const constEtiketIndx As Integer = 6
Private Const constVormIndx As Integer = 7
Private Const constRouteIndx As Integer = 8
Private Const constSterkteIndx As Integer = 9
Private Const constEenheidIndx As Integer = 10
Private Const constStandDoseIndx As Integer = 11
Private Const constDoseEenheidIndx As Integer = 12
Private Const constIndicatiesIndx As Integer = 13
Private Const constFreqIndx As Integer = 14
Private Const constPICU_DoseIndx As Integer = 15
Private Const constPICU_OnderIndx As Integer = 16
Private Const constPICU_BovenIndx As Integer = 17
Private Const constNICU_DoseIndx As Integer = 18
Private Const constNICU_OnderIndx As Integer = 19
Private Const constNICU_BovenIndx As Integer = 20
Private Const constMaxDoseIndx As Integer = 21
Private Const constMaxConcIndx As Integer = 22
Private Const constOplVlstIndx As Integer = 23
Private Const constMinTijdIndx As Integer = 24


Public Sub Formularium_Initialize()

    Dim intN As Integer
    Dim intC As Integer
    Dim strTitle As String

    If Not m_Formularium Is Nothing Then Exit Sub

    strTitle = "Formularium wordt geladen, een ogenblik geduld a.u.b. ..."
    
    ModProgress.StartProgress strTitle
       
    Set m_Formularium = New ClassFormularium
    m_Formularium.GetMedicamenten (True)
    
      
    ModProgress.FinishProgress

End Sub

Public Function Formularium_GetFormularium() As ClassFormularium

    Formularium_Initialize
    Set Formularium_GetFormularium = m_Formularium

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
    
    Dim objMed As ClassMedicatieDisc
    
    On Error GoTo GetMedicamentenError
    
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
            .TherapieGroep = objFormRange.Cells(intN, constHoofdGroepIndx).Value2
            .TherapieSubgroep = objFormRange.Cells(intN, constSubGroepIndx).Value2
            
            .ATC = objFormRange.Cells(intN, constATCIndx).Value2
            .Generiek = objFormRange.Cells(intN, constGeneriekIndx).Value2
            .Vorm = objFormRange.Cells(intN, constVormIndx).Value2
            .Sterkte = objFormRange.Cells(intN, constSterkteIndx).Value2
            .SterkteEenheid = objFormRange.Cells(intN, constEenheidIndx).Value2
            .Etiket = objFormRange.Cells(intN, constEtiketIndx).Value2
            .DeelDose = objFormRange.Cells(intN, constStandDoseIndx).Value2
            .DoseEenheid = objFormRange.Cells(intN, constDoseEenheidIndx).Value2
            
            .SetRouteList objFormRange.Cells(intN, constRouteIndx).Value2
            .SetIndicatieList objFormRange.Cells(intN, constIndicatiesIndx).Value2
            .SetFreqList objFormRange.Cells(intN, constFreqIndx).Value2
            
            .NormDose = objFormRange.Cells(intN, IIf(IsPed, constPICU_DoseIndx, constNICU_DoseIndx)).Value2
            .MinDose = objFormRange.Cells(intN, IIf(IsPed, constPICU_OnderIndx, constNICU_OnderIndx)).Value2
            .MaxDose = objFormRange.Cells(intN, IIf(IsPed, constPICU_BovenIndx, constNICU_BovenIndx)).Value2
            .AbsDose = objFormRange.Cells(intN, constMaxDoseIndx).Value2
            .MaxConc = objFormRange.Cells(intN, constMaxConcIndx).Value2
            .OplVlst = objFormRange.Cells(intN, constOplVlstIndx).Value2
            .MinTijd = objFormRange.Cells(intN, constMinTijdIndx).Value2
        End With
        
        objFormularium.AddMedicament objMed
        Set objMed = Nothing
        
        If blnShowProgress Then ModProgress.SetJobPercentage "Formularium laden", intC, intN
    Next intN
    
    Workbooks(strName).Close

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Exit Sub
    
GetMedicamentenError:
    
    ModLog.LogError "Could not retrieve medicament from: " & strFileName
    
    On Error Resume Next
    
    Workbooks(strName).Close

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    ModProgress.FinishProgress
    
End Sub
