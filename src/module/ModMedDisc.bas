Attribute VB_Name = "ModMedDisc"
Option Explicit

Private Const constATC = "_Glob_MedDisc_ATC_"
Private Const constDoseUnit = "_Glob_MedDisc_DoseEenh_"
Private Const constDoseQty = "_Glob_MedDisc_DoseHoev_"
Private Const constLabel = "_Glob_MedDisc_Etiket_"
Private Const constGeneric = "_Glob_MedDisc_Generic_"
Private Const constGPK = "_Glob_MedDisc_GPK_"
Private Const constIndic = "_Glob_MedDisc_Ind_"
Private Const constTime = "_Glob_MedDisc_Inloop_"
Private Const constDrug = "_Glob_MedDisc_Keuze_"
Private Const constSolNo = "_Glob_MedDisc_OplKeuze_"
Private Const constSolVol = "_Glob_MedDisc_OplVol_"
Private Const constText = "_Glob_MedDisc_Opm_"
Private Const constStandDose = "_Glob_MedDisc_StandDose_"
Private Const constConc = "_Glob_MedDisc_Sterkte_"
Private Const constConcUnit = "_Glob_MedDisc_SterkteEenh_"
Private Const constFreq = "_Glob_MedDisc_Tijden_"
Private Const constRoute = "_Glob_MedDisc_Toed_"

Private Sub MedicamentInvoeren(intN)

    Dim frmMedicament As New FormMedicament
    Dim strMed As String
    Dim strGeneric As String
    Dim strN As String
    
    strN = IIf(intN < 10, "0" & intN, intN)
    
    With frmMedicament
        
        If ModRange.GetRangeValue(constGPK & strN, 0) > 0 Then
            .LoadGPK CStr(ModRange.GetRangeValue(constGPK & strN, vbNullString))
        Else
            .cboGeneriek.Text = ModRange.GetRangeValue(constGeneric & strN, vbNullString)
            .txtSterkte = vbNullString
            .txtSterkteEenheid = vbNullString
        End If
        
        .txtDosisEenheid = ModRange.GetRangeValue(constDoseUnit & strN, vbNullString)
        .txtDosis = ModRange.GetRangeValue(constStandDose & strN, vbNullString)
        .cboRoute = ModRange.GetRangeValue(constRoute & strN, vbNullString)
        .Show
        
        If .lblCancel.Caption = "OK" Then
            strMed = .lblEtiket.Caption
            strGeneric = .cboGeneriek.Text
            
            If strMed = vbNullString And .txtSterkte <> vbNullString Then
                strMed = strGeneric & " " & .txtSterkte & " " & .txtSterkteEenheid
            End If
            
            ModRange.SetRangeValue constDrug & strN, strMed
            ModRange.SetRangeValue constGeneric & strN, strGeneric
            ModRange.SetRangeValue constStandDose & strN, Val(Replace(.txtDosis.Value, ",", "."))
            ModRange.SetRangeValue constDoseUnit & strN, .txtDosisEenheid.Text
            ModRange.SetRangeValue constRoute & strN, .cboRoute.Text
            ModRange.SetRangeValue constGPK & strN, CLng(.GetGPK())

        Else
            If .lblCancel.Caption = "Clear" Then
            
                ModRange.SetRangeValue constDrug & strN, vbNullString
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
                
            End If
        End If
    End With
    
    Set frmMedicament = Nothing

End Sub

Public Sub Medicament_16()

    MedicamentInvoeren 16

End Sub

Public Sub Medicament_17()

    MedicamentInvoeren 17

End Sub

Public Sub Medicament_18()

    MedicamentInvoeren 18

End Sub

Public Sub Medicament_19()

    MedicamentInvoeren 19

End Sub

Public Sub Medicament_15()

    MedicamentInvoeren 15

End Sub

Public Sub Medicament_14()

    MedicamentInvoeren 14

End Sub

Public Sub Medicament_13()

    MedicamentInvoeren 13

End Sub

Public Sub Medicament_12()

    MedicamentInvoeren 12

End Sub

Public Sub Medicament_11()

    MedicamentInvoeren 11

End Sub

Public Sub Medicament_10()

    MedicamentInvoeren 10

End Sub

Public Sub Medicament_9()

    MedicamentInvoeren 9

End Sub

Public Sub Medicament_8()

    MedicamentInvoeren 8

End Sub

Public Sub Medicament_7()

    MedicamentInvoeren 7

End Sub

Public Sub Medicament_6()

    MedicamentInvoeren 6

End Sub

Public Sub Medicament_5()

    MedicamentInvoeren 5

End Sub

Public Sub Medicament_4()

    MedicamentInvoeren 4

End Sub

Public Sub Medicament_3()

    MedicamentInvoeren 3

End Sub

Public Sub Medicament_2()

    MedicamentInvoeren 2

End Sub

Public Sub Medicament_1()

    MedicamentInvoeren 1

End Sub


Public Sub Medicament_20()

    MedicamentInvoeren 20

End Sub

Public Sub Medicament_21()

    MedicamentInvoeren 21

End Sub

Public Sub Medicament_22()

    MedicamentInvoeren 22

End Sub

Public Sub Medicament_23()

    MedicamentInvoeren 23

End Sub

Public Sub Medicament_24()

    MedicamentInvoeren 24

End Sub

Public Sub Medicament_25()

    MedicamentInvoeren 25

End Sub

Public Sub Medicament_26()

    MedicamentInvoeren 26

End Sub

Public Sub Medicament_27()

    MedicamentInvoeren 27

End Sub

Public Sub Medicament_28()

    MedicamentInvoeren 28

End Sub

Public Sub Medicament_29()

    MedicamentInvoeren 29

End Sub

Public Sub Medicament_30()

    MedicamentInvoeren 30

End Sub

Private Sub OpmMedDisc(intN As Integer)
    
    Dim frmOpmerking As New FormOpmerking
    Dim strRange As String
    
    strRange = shtGlobBerOpm.Name & "!c" & intN

    frmOpmerking.txtOpmerking.Text = Range(strRange).Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        ModRange.SetRangeValue strRange, frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString
    
    Set frmOpmerking = Nothing

End Sub

Public Sub OpmMedDisc_1()
    
    OpmMedDisc 16
    
End Sub

Public Sub OpmMedDisc_2()
    
    OpmMedDisc 17

End Sub

Public Sub OpmMedDisc_3()
    
    OpmMedDisc 18

End Sub
Public Sub OpmMedDisc_4()
    
    OpmMedDisc 19

End Sub

Public Sub OpmMedDisc_5()
    
    OpmMedDisc 20

End Sub
Public Sub OpmMedDisc_6()
    
    OpmMedDisc 21

End Sub
Public Sub OpmMedDisc_7()
    
    OpmMedDisc 22

End Sub
Public Sub OpmMedDisc_8()
    
    OpmMedDisc 23

End Sub
Public Sub OpmMedDisc_9()
    
    OpmMedDisc 24

End Sub
Public Sub OpmMedDisc_10()
    
    OpmMedDisc 25

End Sub
Public Sub OpmMedDisc_11()
    
    OpmMedDisc 26

End Sub
Public Sub OpmMedDisc_12()
    
    OpmMedDisc 27

End Sub
Public Sub OpmMedDisc_13()
    
    OpmMedDisc 28

End Sub
Public Sub OpmMedDisc_14()
    
    OpmMedDisc 29

End Sub
Public Sub OpmMedDisc_15()
    
    OpmMedDisc 30

End Sub
Public Sub OpmMedDisc_16()
    
    OpmMedDisc 31

End Sub
Public Sub OpmMedDisc_17()
    
    OpmMedDisc 32

End Sub
Public Sub OpmMedDisc_18()
    
    OpmMedDisc 33

End Sub
Public Sub OpmMedDisc_19()
    
    OpmMedDisc 34

End Sub
Public Sub OpmMedDisc_20()
    
    OpmMedDisc 35

End Sub

Public Sub OpmMedDisc_21()
    
    OpmMedDisc 36

End Sub

Public Sub OpmMedDisc_22()
    
    OpmMedDisc 37

End Sub

Public Sub OpmMedDisc_23()
    
    OpmMedDisc 38

End Sub

Public Sub OpmMedDisc_24()
    
    OpmMedDisc 39

End Sub

Public Sub OpmMedDisc_25()
    
    OpmMedDisc 40

End Sub

Public Sub OpmMedDisc_26()
    
    OpmMedDisc 41

End Sub

Public Sub OpmMedDisc_27()
    
    OpmMedDisc 42

End Sub

Public Sub OpmMedDisc_28()
    
    OpmMedDisc 43

End Sub

Public Sub OpmMedDisc_29()
    
    OpmMedDisc 44

End Sub

Public Sub OpmMedDisc_30()
    
    OpmMedDisc 45

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

