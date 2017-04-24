Attribute VB_Name = "ModNeoPrint"
Option Explicit

Private Const constDrugNo As String = "Var_Neo_PrintApothNo"
Private Const constMedKeuze As String = "Var_Neo_InfB_Cont_MedKeuze_"

Private Sub PrintSheet(ByRef shtSheet As Worksheet)

    shtSheet.Unprotect ModConst.CONST_PASSWORD
    shtSheet.PrintPreview
    If Not ModSetting.GetDevelopmentMode Then shtSheet.Protect ModConst.CONST_PASSWORD
    
End Sub

Private Sub PrintBriefNo(ByVal intNo As Integer)

    ModRange.SetRangeValue constDrugNo, intNo
    
    PrintSheet shtNeoPrtApoth
    
End Sub

Private Sub TestPrintApotheekBrief()

    PrintBriefNo 1

End Sub

Public Sub PrintApotheekWerkBrief()

    Dim intNo As Integer
    Dim intMed As Integer
    Dim strNo As String
    
    For intNo = 1 To 10
    
        strNo = IIf(intNo < 10, "0" & intNo, intNo)
        intMed = ModRange.GetRangeValue(constMedKeuze & strNo, 0)
        
        If intMed > 1 Then
            PrintBriefNo intNo
        End If
    
    Next intNo

End Sub
