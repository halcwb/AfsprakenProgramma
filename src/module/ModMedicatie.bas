Attribute VB_Name = "ModMedicatie"
Option Explicit

Private Const constTblNeoMedCont As String = "Tbl_Neo_MedIV"

Public Function Medicatie_CalcEpiQty(ByVal dblWght As Double) As Double

    Dim dblQty As Double
    Dim dblMultiple As Double
    Dim intFactor As Integer
    
    dblQty = IIf(dblWght >= 25, dblWght, 2 * dblWght)
    dblQty = IIf(dblWght >= 48, 48, dblQty)
    
    dblMultiple = 0.1 'ModExcel.Excel_Index(constTblMedIV, intMed, 4)
    intFactor = 1                                                   ' >=10ml  hele getallen geen decimalen
    intFactor = IIf(dblQty / dblMultiple < 10, 10, intFactor)       ' >=1,0  <10 ml   1 decimaal nauwkeurig 0,1
    intFactor = IIf(dblQty / dblMultiple < 1, 100, intFactor)       ' >=0,1 < 1,0 ml:     2 decimalen nauwkeurig 0,01
    intFactor = IIf(dblQty / dblMultiple <= 0.1, 100, intFactor)    ' <0,1ml  2 decimalen nauwkeurig 0,01 + verdunningstekst
    dblMultiple = dblMultiple / intFactor
    
    dblQty = ModExcel.Excel_RoundBy(dblQty, dblMultiple)
    If dblQty = 0 Then dblQty = dblMultiple
    
    Medicatie_CalcEpiQty = dblQty

End Function

Public Function Medicatie_CalcEpiVol(ByVal dblWght As Double) As Double

    Medicatie_CalcEpiVol = IIf(dblWght < 6, 24, 48)

End Function

Public Function Medicatie_CalcEpiStand(ByVal dblWght As Double) As Double
    
    Dim dblStand As Double

    dblStand = IIf(dblWght >= 7, 2, 1)
    dblStand = IIf(dblWght >= 25, 4, dblStand)
    
    Medicatie_CalcEpiStand = dblStand

End Function

Public Function Medicatie_IsEpiduraal(ByVal strMed As String) As Boolean

    Medicatie_IsEpiduraal = ModString.ContainsCaseSensitive(strMed, "EPI ") Or ModString.ContainsCaseSensitive(strMed, "Epi ")

End Function

Private Function GetMedContIVName(ByVal intMed As Integer, ByVal strTbl As String) As String

    GetMedContIVName = ModExcel.Excel_Index(strTbl, intMed, 1)

End Function

Public Function GetNeoMedContIVName(ByVal intMed As Integer) As String

    GetNeoMedContIVName = GetMedContIVName(intMed, constTblNeoMedCont)

End Function

Private Sub Test_GetNeoMedContIVName()

    MsgBox GetNeoMedContIVName(2)

End Sub
