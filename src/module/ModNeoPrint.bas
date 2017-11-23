Attribute VB_Name = "ModNeoPrint"
Option Explicit

Private Const constDrugNo As String = "Var_Neo_PrintApothNo"
Private Const constMedKeuze As String = "Var_Neo_InfB_Cont_MedKeuze_"
Private Const constHospNo As String = "__0_PatNum"

Private Sub PrintBriefNo(ByVal intNo As Integer, ByVal blnAsk As Boolean, ByVal blnPrev As Boolean, ByVal strFile As String)

    shtNeoPrtApoth.Unprotect ModConst.CONST_PASSWORD
    ModRange.SetRangeValue constDrugNo, intNo
    
    If strFile = vbNullString Then
        PrintSheet shtNeoPrtApoth, 2, blnAsk, blnPrev
    Else
        SaveSheetAsPDF shtNeoPrtApoth, strFile & "_" & IntNToStrN(intNo) & "_" & ".pdf"
    End If
End Sub

Private Sub TestPrintApotheekBrief()

    PrintBriefNo 1, False, True, vbNullString

End Sub

Public Sub PrintApotheekWerkBrief()

    If Not ModNeoInfB.NeoInfB_IsValidContMed Then
        ModMessage.ShowMsgBoxExclam "Continue medicatie bevat fouten, kan de apotheek bereidingsvoorschriften niet afdrukken"
        Exit Sub
    End If

    PrintApotheekWerkBriefPrev True, vbNullString

End Sub

Public Sub SaveApotheekWerkBrief(ByVal strPDF As String)

    PrintApotheekWerkBriefPrev False, strPDF

End Sub

Private Sub PrintApotheekWerkBriefPrev(ByVal blnPrev As Boolean, ByVal strFile As String)

    Dim intNo As Integer
    Dim intMed As Integer
    Dim strNo As String
    Dim blnAsk As Boolean
    Dim blnPrint As Boolean
    Dim vbAnswer As Integer
    
    blnPrint = True
    blnPrint = Not ModRange.GetRangeValue(constHospNo, vbNullString) = vbNullString
    If Not NeoInfB_Is1700() And blnPrint Then
        vbAnswer = ModMessage.ShowMsgBoxYesNo("Huidige infuusbrief is niet de 17:00 versie!" & vbNewLine & "Toch printen?")
        blnPrint = blnPrint And vbAnswer = vbYes
    End If
    
    If Not blnPrint Then Exit Sub
    
    blnAsk = False
    For intNo = 1 To 10
    
        strNo = IIf(intNo < 10, "0" & intNo, intNo)
        intMed = ModRange.GetRangeValue(constMedKeuze & strNo, 0)
        
        If intMed > 1 Then
            PrintBriefNo intNo, blnAsk, blnPrev, strFile
            blnAsk = False
        End If
    
    Next intNo

End Sub

Public Sub PrintNeoWerkBrief()

    If Not ModNeoInfB.NeoInfB_IsValidContMed Then
        ModMessage.ShowMsgBoxExclam "Continue medicatie bevat fouten, kan de werkbrief niet afdrukken"
        Exit Sub
    End If
    
    If Not ModNeoInfB.NeoInfB_IsValidTPN Then
        ModMessage.ShowMsgBoxExclam "TPN bevat fouten, kan de werkbrief niet afdrukken"
        Exit Sub
    End If
    
    PrintNeoWerkBriefPrev True, vbNullString

End Sub

Public Sub SaveNeoWerkBrief(ByVal strFile As String)

    PrintNeoWerkBriefPrev False, strFile

End Sub

Private Sub PrintNeoWerkBriefPrev(ByVal blnPrev As Boolean, ByVal strFile As String)

    shtNeoPrtWerkbr.Unprotect ModConst.CONST_PASSWORD
    
    If strFile = vbNullString Then
        PrintSheet shtNeoPrtWerkbr, 1, False, blnPrev
    Else
        SaveSheetAsPDF shtNeoPrtWerkbr, strFile & ".pdf"
    End If
    
End Sub

