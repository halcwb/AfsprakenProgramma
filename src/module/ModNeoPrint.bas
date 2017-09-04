Attribute VB_Name = "ModNeoPrint"
Option Explicit

Private Const constDrugNo As String = "Var_Neo_PrintApothNo"
Private Const constMedKeuze As String = "Var_Neo_InfB_Cont_MedKeuze_"

Private Sub PrintBriefNo(ByVal intNo As Integer, ByVal blnAsk As Boolean)

    shtNeoPrtApoth.Unprotect ModConst.CONST_PASSWORD
    ModRange.SetRangeValue constDrugNo, intNo
    
    PrintSheet shtNeoPrtApoth, 2, blnAsk
    
End Sub

Private Sub TestPrintApotheekBrief()

    PrintBriefNo 1, True

End Sub

Public Sub PrintApotheekWerkBrief()

    Dim intNo As Integer
    Dim intMed As Integer
    Dim strNo As String
    Dim blnAsk As Boolean
    
    blnAsk = True
    For intNo = 1 To 10
    
        strNo = IIf(intNo < 10, "0" & intNo, intNo)
        intMed = ModRange.GetRangeValue(constMedKeuze & strNo, 0)
        
        If intMed > 1 Then
            PrintBriefNo intNo, blnAsk
            blnAsk = False
        End If
    
    Next intNo

End Sub
