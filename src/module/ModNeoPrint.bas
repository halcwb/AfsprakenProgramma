Attribute VB_Name = "ModNeoPrint"
Option Explicit

Private Const constDrugNo As String = "Var_Neo_PrintApothNo"
Private Const constMedKeuze As String = "Var_Neo_InfB_Cont_MedKeuze_"
Private Const constHospNo As String = "__0_PatNum"
Private Const constUserType As String = "_User_Type"


Private Function PrintBriefNo(ByVal intNo As Integer, ByVal blnAsk As Boolean, ByVal blnPrev As Boolean, ByVal strFile As String) As String
    
    Dim strPDF As String

    shtNeoPrtApoth.Unprotect ModConst.CONST_PASSWORD
    ModRange.SetRangeValue constDrugNo, intNo
    
    strPDF = vbNullString
    If strFile = vbNullString Then
        PrintSheet shtNeoPrtApoth, 2, blnAsk, blnPrev
    Else
        strPDF = strFile & "_" & IntNToStrN(intNo) & "_" & ".pdf"
        SaveSheetAsPDF shtNeoPrtApoth, strFile & "_" & IntNToStrN(intNo) & "_" & ".pdf", False
    End If
    
    PrintBriefNo = strPDF
    
End Function

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

Public Sub SendApotheekWerkBrief()

    Dim intNo As Integer
    Dim intMed As Integer
    Dim strNo As String
    Dim blnAsk As Boolean
    Dim blnPrint As Boolean
    Dim strUser As String
    Dim vbAnswer As Integer
    
    Dim objMsg As Object
    Dim strTo As String
    Dim strFrom As String
    Dim strSubject As String
    Dim strHTML As String
    
    Dim strFile As String
    Dim strPDF As String
    Dim strPDFList As String
    Dim arrPDF() As String
    Dim intC As Integer
    Dim intN As Integer
    
    Dim strMail As String
    
    On Error GoTo SendApotheekWerkBriefError
       
    If Not ModNeoInfB.NeoInfB_IsValidContMed Then
        ModProgress.FinishProgress
        ModMessage.ShowMsgBoxExclam "Continue medicatie bevat fouten." & "Kan de apotheek bereidingsvoorschriften niet verzenden!"
        Exit Sub
    End If
    
    blnPrint = Not ModRange.GetRangeValue(constHospNo, vbNullString) = vbNullString
    If Not blnPrint Then
        ModMessage.ShowMsgBoxExclam "Er is geen ziekenhuis nummer opgegeven." & vbNewLine & "Kan de apotheekbrief niet verzenden!"
        Exit Sub
    End If
    
    strUser = ModRange.GetRangeValue(constUserType, vbNullString)
    If Not (strUser = "Supervisor" Or strUser = "Artsen") And ModSetting.IsProductionDir() Then
        ModMessage.ShowMsgBoxExclam "Er is geen arts ingelogd." & vbNewLine & "Kan de apotheekbrief niet verzenden!"
        Exit Sub
    End If
           
    strMail = "wkz-algemeen@umcutrecht.nl"
    If Not ModSetting.IsProductionDir() Then strMail = ModMessage.ShowInputBox("Voer een email adres in", vbNullString)
    
    If strMail = vbNullString Then
        ModMessage.ShowMsgBoxExclam "Er moet een email adres worden ingevoerd." & vbNewLine & "Kan de apotheekbrief niet verzenden!"
        Exit Sub
    End If
    
    ModBed.CloseBed False
    
    ModProgress.StartProgress "Medicatie naar de apotheek verzenden"

    ModNeoInfB.NeoInfB_SelectInfB True, False
    
    strTo = strMail
    strFrom = "FunctioneelBeheerMetavision@umcutrecht.nl"
    strSubject = "NICU VTGM protocollen voor " & ModPatient.PatientHospNum & " " & ModPatient.PatientAchterNaam & ", " & ModPatient.PatientVoorNaam
    strHTML = vbNullString
    
    Set objMsg = CreateObject("CDO.Message")
    With objMsg
         
        .to = CStr(strTo)
        .From = CStr(strFrom)
        .Subject = CStr(strSubject)
        .HTMLBody = CStr(strHTML)
        .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'cdoSendUsingPickup=1, cdoSendUsingPort=2, cdoSendUsingExchange=3
        .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.umcutrecht.nl"
        .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        .Configuration.Fields.Update
        
        strFile = Environ("TEMP") & "\Werkbrief_" & ModPatient.PatientHospNum
        strPDF = PrintNeoWerkBriefPrev(False, strFile)
        .AddAttachment strPDF
        strPDFList = strPDFList & strPDF & vbNewLine
        
        For intNo = 1 To 10

            strNo = IIf(intNo < 10, "0" & intNo, intNo)
            intMed = ModRange.GetRangeValue(constMedKeuze & strNo, 0)

            If intMed > 1 Then
                strFile = Environ("TEMP") & "\VTGM_" & ModPatient.PatientHospNum
                strPDF = PrintBriefNo(intNo, False, False, strFile)
                .AddAttachment strPDF
                strPDFList = strPDFList & strPDF & vbNewLine

            End If

            ModProgress.SetJobPercentage "Verzenden", 10, intNo

        Next intNo
        
        .Send
    
    End With
    
    arrPDF = Split(strPDFList, vbNewLine)
    intC = UBound(arrPDF)
    For intN = 0 To intC
        strPDF = Trim(arrPDF(intN))
        If Not strPDF = vbNullString Then Kill strPDF
    Next
    
    Set objMsg = Nothing
    
    ModProgress.FinishProgress
    
    ModMessage.ShowMsgBoxInfo "Medicatie is verstuurd naar de apotheek en de patient is opgeslagen"
    
    Exit Sub

SendApotheekWerkBriefError:

    ModLog.LogError "SendApotheekWerkBriefError"
    
    On Error Resume Next
    
    ModMessage.ShowMsgBoxError "TPN is niet verstuurd naar de apotheek, een foutmelding is verzonden naar functioneel beheer"
    
    arrPDF = Split(strPDFList, vbNewLine)
    intC = UBound(arrPDF)
    For intN = 0 To intC
        strPDF = Trim(arrPDF(intN))
        If Not strPDF = vbNullString Then Kill strPDF
    Next
    
    Set objMsg = Nothing
    
    ModProgress.FinishProgress

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
        ModMessage.ShowMsgBoxExclam "Huidige infuusbrief is niet de 17:00 versie!" & vbNewLine & "Kan de apotheekbrief niet printen!"
        Exit Sub
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

    If Not NeoInfB_Is1700() Then
        ModMessage.ShowMsgBoxExclam "Huidige infuusbrief is niet de 17:00 versie!" & vbNewLine & "Kan de werkbrief niet printen!"
        Exit Sub
    End If

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

Private Function PrintNeoWerkBriefPrev(ByVal blnPrev As Boolean, ByVal strFile As String) As String

    Dim strPDF As String

    shtNeoPrtWerkbr.Unprotect ModConst.CONST_PASSWORD
    
    If strFile = vbNullString Then
        PrintSheet shtNeoPrtWerkbr, 1, False, blnPrev
    Else
        strPDF = strFile & ".pdf"
        SaveSheetAsPDF shtNeoPrtWerkbr, strPDF, True
    End If
    
    PrintNeoWerkBriefPrev = strPDF
    
End Function

Public Sub NeoPrint_PrintMedicatieDisc(ByVal blnPrev As Boolean)

    ModSheet.PrintSheet shtNeoPrtMedDisc, 1, False, True

End Sub

Public Sub NeoPrint_PrintMedicatieCont(ByVal blnPrev As Boolean)

    ModSheet.PrintSheetAllPortrait shtNeoPrtAfspr
    ModSheet.PrintSheet shtNeoPrtAfspr, 1, False, True

End Sub

Public Sub NeoPrint_PrintAcuteBlad(ByVal blnPrev As Boolean)

    ModSheet.PrintSheetAllPortrait shtNeoGuiAcuut
    ModSheet.PrintSheet shtNeoGuiAcuut, 1, False, True

End Sub

