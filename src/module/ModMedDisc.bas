Attribute VB_Name = "ModMedDisc"
Option Explicit

Private Const constFormularium = "Formularium.xlsx"

Private Const constATC As String = "_Glob_MedDisc_ATC_"
Private Const constDoseUnit As String = "_Glob_MedDisc_DoseEenh_"
Private Const constDoseQty As String = "_Glob_MedDisc_DoseHoev_"
Private Const constLabel As String = "_Glob_MedDisc_Etiket_"
Private Const constGeneric As String = "_Glob_MedDisc_Generic_"
Private Const constGPK As String = "_Glob_MedDisc_GPK_"
Private Const constIndic As String = "_Glob_MedDisc_Ind_"
Private Const constTime As String = "_Glob_MedDisc_Inloop_"
Private Const constDrug As String = "_Glob_MedDisc_Keuze_"
Private Const constSolNo As String = "_Glob_MedDisc_OplKeuze_"
Private Const constSolVol As String = "_Glob_MedDisc_OplVol_"
Private Const constText As String = "_Glob_MedDisc_Opm_"
Private Const constStandDose As String = "_Glob_MedDisc_StandDose_"
Private Const constConc As String = "_Glob_MedDisc_Sterkte_"
Private Const constConcUnit As String = "_Glob_MedDisc_SterkteEenh_"
Private Const constFreq As String = "_Glob_MedDisc_Tijden_"
Private Const constRoute As String = "_Glob_MedDisc_Toed_"

Private Sub Clear(ByVal intN As Integer)

    Dim strN As String
    
    strN = IIf(intN < 10, "0" & intN, intN)

    ModRange.SetRangeValue constLabel, vbNullString
    ModRange.SetRangeValue constIndic, vbNullString
    ModRange.SetRangeValue constATC & strN, vbNullString
    ModRange.SetRangeValue constDrug & strN, vbNullString
    ModRange.SetRangeValue constStandDose & strN, vbNullString
    ModRange.SetRangeValue constDoseUnit & strN, vbNullString
    ModRange.SetRangeValue constRoute & strN, vbNullString
    ModRange.SetRangeValue constText & strN, vbNullString
    ModRange.SetRangeValue constDoseQty & strN, vbNullString
    ModRange.SetRangeValue constConc, 0
    ModRange.SetRangeValue constConcUnit, vbNullString
    ModRange.SetRangeValue constFreq & strN, 1
    ModRange.SetRangeValue constSolVol & strN, 0
    ModRange.SetRangeValue constSolNo & strN, 0
    ModRange.SetRangeValue constTime & strN, 0
    ModRange.SetRangeValue constGPK & strN, 0

End Sub

Public Sub MedDisc_Clear_01()

    Clear 1

End Sub

Public Sub GetMedicamenten(ByRef objFormularium As ClassFormularium, ByVal blnShowProgress As Boolean)

    Dim intN As Integer
    Dim intC As Integer
    Dim objFormRange As Range
    Dim objSheet As Worksheet
    
    Dim strFileName As String
    Dim strName As String
    Dim strSheet As String
    
    Dim strTitle As String
    
    On Error GoTo GetMedicamentenError
    
    strName = "Formularium.xlsx"
    strSheet = "Table"
    
    Application.DisplayAlerts = False

    strFileName = ModMedDisc.GetFormulariumDatabasePath() + strName

    Workbooks.Open strFileName, True, True
    
    Set objSheet = Workbooks(strName).Worksheets(strSheet)
    Set objFormRange = objSheet.Range("A1").CurrentRegion
        
    intC = objFormRange.Rows.Count
    For intN = 2 To intC
        objFormularium.AddMedicament (objFormRange.Cells(intN, 1))
        With objFormularium.Item(intN - 1)
            .ATC = objFormRange.Cells(intN, 2)
            .TherapieGroep = objFormRange.Cells(intN, 3)
            .TherapieSubgroep = objFormRange.Cells(intN, 4)
            .Generiek = objFormRange.Cells(intN, 5)
            .Etiket = objFormRange.Cells(intN, 6)
            .Vorm = objFormRange.Cells(intN, 7)
            .routes = objFormRange.Cells(intN, 8)
            .Sterkte = objFormRange.Cells(intN, 9)
            .SterkteEenheid = objFormRange.Cells(intN, 10)
            .Dosis = objFormRange.Cells(intN, 11)
            .DosisEenheid = objFormRange.Cells(intN, 12)
            .Indicaties = objFormRange.Cells(intN, 13)
        End With
        
        If blnShowProgress Then ModProgress.SetJobPercentage "Formularium laden", intC, intN
    Next intN
    
    Workbooks(strName).Close
    Application.DisplayAlerts = True
    
    Exit Sub
    
GetMedicamentenError:
    
    ModLog.LogError "Could not retrieve medicament from: " & strFileName
    
End Sub

Private Sub MedicamentInvoeren(ByVal intN As Integer)

    Dim objMed As ClassMedicatieDisc
    Dim strMed As String
    Dim strGeneric As String
    Dim strN As String
      
    strN = IIf(intN < 10, "0" & intN, intN)
    
    With FormMedicament
        
        If ModRange.GetRangeValue(constGPK & strN, 0) > 0 Then
            .LoadGPK CStr(ModRange.GetRangeValue(constGPK & strN, vbNullString))
        Else
            .cboGeneriek.Text = ModRange.GetRangeValue(constGeneric & strN, vbNullString)
            .txtSterkte.Text = vbNullString
            .txtSterkteEenheid.Text = vbNullString
        End If
        
        .txtDosisEenheid.Text = ModRange.GetRangeValue(constDoseUnit & strN, vbNullString)
        .txtDosis.Text = ModRange.GetRangeValue(constStandDose & strN, vbNullString)
        .cboRoute.Text = ModRange.GetRangeValue(constRoute & strN, vbNullString)
        .Show
        
        If .GetClickedButton = "OK" Then
            If .HasSelectedMedicament() Then
                Set objMed = .GetSelectedMedicament()
                strMed = objMed.GetMedicamentText()
                strGeneric = objMed.Generiek
                
                ModRange.SetRangeValue constDrug & strN, strMed
                ModRange.SetRangeValue constGeneric & strN, strGeneric
                ModRange.SetRangeValue constStandDose & strN, Val(Replace(objMed.Dosis, ",", "."))
                ModRange.SetRangeValue constDoseUnit & strN, objMed.DosisEenheid
                ModRange.SetRangeValue constRoute & strN, .GetSelectedRoute()
                ModRange.SetRangeValue constIndic & strN, .GetSelectedIndication()
                ModRange.SetRangeValue constGPK & strN, CLng(objMed.GPK)
                ModRange.SetRangeValue constATC & strN, objMed.ATC
                ModRange.SetRangeValue constConc & strN, Conversion.CDbl(objMed.Sterkte)
                ModRange.SetRangeValue constConcUnit & strN, objMed.SterkteEenheid
                ModRange.SetRangeValue constLabel & strN, objMed.Etiket
            End If

        Else
            If .GetClickedButton = "Clear" Then

                ModRange.SetRangeValue constDrug & strN, vbNullString
                ModRange.SetRangeValue constGeneric & strN, vbNullString
                ModRange.SetRangeValue constStandDose & strN, vbNullString
                ModRange.SetRangeValue constDoseUnit & strN, vbNullString
                ModRange.SetRangeValue constRoute & strN, vbNullString
                ModRange.SetRangeValue constText & strN, vbNullString
                ModRange.SetRangeValue constDoseQty & strN, vbNullString
                ModRange.SetRangeValue constFreq & strN, 1
                ModRange.SetRangeValue constSolVol & strN, 0
                ModRange.SetRangeValue constSolNo & strN, 0
                ModRange.SetRangeValue constTime & strN, 0
                ModRange.SetRangeValue constGPK & strN, 0
                ModRange.SetRangeValue constATC & strN, vbNullString
                ModRange.SetRangeValue constConc & strN, 0
                ModRange.SetRangeValue constConcUnit & strN, vbNullString
                ModRange.SetRangeValue constLabel & strN, vbNullString

            End If
        End If
    End With

End Sub

Public Sub MedDisc_EnterMed_01()

    MedicamentInvoeren 1

End Sub

Public Sub MedDisc_EnterMed_02()

    MedicamentInvoeren 2

End Sub

Public Sub MedDisc_EnterMed_03()

    MedicamentInvoeren 3

End Sub

Public Sub MedDisc_EnterMed_04()

    MedicamentInvoeren 4

End Sub

Public Sub MedDisc_EnterMed_05()

    MedicamentInvoeren 5

End Sub

Public Sub MedDisc_EnterMed_06()

    MedicamentInvoeren 6

End Sub

Public Sub MedDisc_EnterMed_07()

    MedicamentInvoeren 7

End Sub

Public Sub MedDisc_EnterMed_08()

    MedicamentInvoeren 8

End Sub

Public Sub MedDisc_EnterMed_09()

    MedicamentInvoeren 9

End Sub

Public Sub MedDisc_EnterMed_10()

    MedicamentInvoeren 10

End Sub

Public Sub MedDisc_EnterMed_11()

    MedicamentInvoeren 11

End Sub

Public Sub MedDisc_EnterMed_12()

    MedicamentInvoeren 12

End Sub

Public Sub MedDisc_EnterMed_13()

    MedicamentInvoeren 13

End Sub

Public Sub MedDisc_EnterMed_14()

    MedicamentInvoeren 14

End Sub

Public Sub MedDisc_EnterMed_15()

    MedicamentInvoeren 15

End Sub

Public Sub MedDisc_EnterMed_16()

    MedicamentInvoeren 16

End Sub

Public Sub MedDisc_EnterMed_17()

    MedicamentInvoeren 17

End Sub

Public Sub MedDisc_EnterMed_18()

    MedicamentInvoeren 18

End Sub

Public Sub MedDisc_EnterMed_19()

    MedicamentInvoeren 19

End Sub

Public Sub MedDisc_EnterMed_20()

    MedicamentInvoeren 20

End Sub

Public Sub MedDisc_EnterMed_21()

    MedicamentInvoeren 21

End Sub

Public Sub MedDisc_EnterMed_22()

    MedicamentInvoeren 22

End Sub

Public Sub MedDisc_EnterMed_23()

    MedicamentInvoeren 23

End Sub

Public Sub MedDisc_EnterMed_24()

    MedicamentInvoeren 24

End Sub

Public Sub MedDisc_EnterMed_25()

    MedicamentInvoeren 25

End Sub

Public Sub MedDisc_EnterMed_26()

    MedicamentInvoeren 26

End Sub

Public Sub MedDisc_EnterMed_27()

    MedicamentInvoeren 27

End Sub

Public Sub MedDisc_EnterMed_28()

    MedicamentInvoeren 28

End Sub

Public Sub MedDisc_EnterMed_29()

    MedicamentInvoeren 29

End Sub

Public Sub MedDisc_EnterMed_30()

    MedicamentInvoeren 30

End Sub

Private Sub OpmMedDisc(ByVal intN As Integer)
    
    Dim frmOpmerking As FormOpmerking
    Dim strRange As String
    
    Set frmOpmerking = New FormOpmerking
    
    strRange = constText
    strRange = constText & IIf(intN < 10, "0" & intN, intN)

    frmOpmerking.txtOpmerking.Text = Range(strRange).Value
    frmOpmerking.Show
    
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        ModRange.SetRangeValue strRange, frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString
    
    Set frmOpmerking = Nothing

End Sub

Public Sub MedDisc_EnterText_01()
    
    OpmMedDisc 1
    
End Sub

Public Sub MedDisc_EnterText_02()
    
    OpmMedDisc 2

End Sub

Public Sub MedDisc_EnterText_03()
    
    OpmMedDisc 3

End Sub
Public Sub MedDisc_EnterText_04()
    
    OpmMedDisc 4

End Sub

Public Sub MedDisc_EnterText_05()
    
    OpmMedDisc 5

End Sub

Public Sub MedDisc_EnterText_06()
    
    OpmMedDisc 6

End Sub

Public Sub MedDisc_EnterText_07()
    
    OpmMedDisc 7

End Sub

Public Sub MedDisc_EnterText_08()
    
    OpmMedDisc 8

End Sub

Public Sub MedDisc_EnterText_09()
    
    OpmMedDisc 9

End Sub

Public Sub MedDisc_EnterText_10()
    
    OpmMedDisc 10

End Sub

Public Sub MedDisc_EnterText_11()
    
    OpmMedDisc 11

End Sub

Public Sub MedDisc_EnterText_12()
    
    OpmMedDisc 12

End Sub

Public Sub MedDisc_EnterText_13()
    
    OpmMedDisc 13

End Sub

Public Sub MedDisc_EnterText_14()
    
    OpmMedDisc 14

End Sub

Public Sub MedDisc_EnterText_15()
    
    OpmMedDisc 15

End Sub

Public Sub MedDisc_EnterText_16()
    
    OpmMedDisc 16

End Sub

Public Sub MedDisc_EnterText_17()
    
    OpmMedDisc 17

End Sub

Public Sub MedDisc_EnterText_18()
    
    OpmMedDisc 18

End Sub

Public Sub MedDisc_EnterText_19()
    
    OpmMedDisc 19

End Sub

Public Sub MedDisc_EnterText_20()
    
    OpmMedDisc 20

End Sub

Public Sub MedDisc_EnterText_21()
    
    OpmMedDisc 21

End Sub

Public Sub MedDisc_EnterText_22()
    
    OpmMedDisc 22

End Sub

Public Sub MedDisc_EnterText_23()
    
    OpmMedDisc 23

End Sub

Public Sub MedDisc_EnterText_24()
    
    OpmMedDisc 24

End Sub

Public Sub MedDisc_EnterText_25()
    
    OpmMedDisc 25

End Sub

Public Sub MedDisc_EnterText_26()
    
    OpmMedDisc 26

End Sub

Public Sub MedDisc_EnterText_27()
    
    OpmMedDisc 27

End Sub

Public Sub MedDisc_EnterText_28()
    
    OpmMedDisc 28

End Sub

Public Sub MedDisc_EnterText_29()
    
    OpmMedDisc 29

End Sub

Public Sub MedDisc_EnterText_30()
    
    OpmMedDisc 30

End Sub

' Make sure that the active workbook is Afspraken2015.xlsm
' and return the path of the Formularium workbook
Public Function GetFormulariumDatabasePath() As String
    Dim strPath As String
    Dim arrPath() As String
    Dim intCounter As Integer

    strPath = vbNullString
    arrPath = Split(WbkAfspraken.Path, "\")
    
    ' create the path 2 dirs down workbook path
    For intCounter = 0 To (UBound(arrPath) - 2)
        strPath = strPath & arrPath(intCounter) & "\"
    Next
    
    GetFormulariumDatabasePath = strPath & "db\"

End Function

