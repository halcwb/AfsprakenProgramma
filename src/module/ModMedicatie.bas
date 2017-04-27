Attribute VB_Name = "ModMedicatie"
Option Explicit

Public Function Medicatie_CalcEpiQty(ByVal dblWght As Double)

    Medicatie_CalcEpiQty = dblWght

End Function

Public Function Medicatie_CalcEpiVol(ByVal dblWght As Double)

    Medicatie_CalcEpiVol = IIf(dblWght < 6, 24, 48)

End Function

Public Function Medicatie_CalcEpiStand(ByVal dblWght As Double)

    Medicatie_CalcEpiStand = IIf(dblWght < 6, 1, 2)

End Function

Private Function GetMedContIVName(ByVal intMed As Integer, ByVal strTbl As String) As String

    GetMedContIVName = ModExcel.Excel_Index(strTbl, intMed, 1)

End Function

Public Function GetNeoMedContIVName(ByVal intMed As Integer)

    GetNeoMedContIVName = GetMedContIVName(intMed, "Tbl_Neo_MedIV")

End Function

Private Sub Test_GetNeoMedContIVName()

    MsgBox GetNeoMedContIVName(2)

End Sub
