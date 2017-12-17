Attribute VB_Name = "ModPedPrint"
Option Explicit

Private Const CONST_TPN_1 As Integer = 2
Private Const CONST_TPN_2 As Integer = 7
Private Const CONST_TPN_3 As Integer = 16
Private Const CONST_TPN_4 As Integer = 30
Private Const CONST_TPN_5 As Integer = 50

Public Sub SaveAndPrintAfspraken()

    Dim frmPrintAfspraken As FormPrintAfspraken
    
    Set frmPrintAfspraken = New FormPrintAfspraken
    ModBed.CloseBed True
    frmPrintAfspraken.Show
    
End Sub

Public Sub PedPrint_PrintMedicatieCont(ByVal blnPrev As Boolean)

    ModSheet.PrintSheet shtPedPrtAfspr, 1, False, True

End Sub

Public Sub PedPrint_PrintMedicatieDisc(ByVal blnPrev As Boolean)

    ModSheet.PrintSheet shtPedPrtMedDisc, 1, False, True

End Sub

Public Sub PedPrint_PrintTPN(ByVal blnPrev As Boolean)

    PrintTPN blnPrev, vbNullString

End Sub

Public Sub PedPrint_PrintAcuteBlad(ByVal blnPrev As Boolean)

    ModSheet.PrintSheetAllPortrait shtPedGuiAcuut
    ModSheet.PrintSheet shtPedGuiAcuut, 1, False, True

End Sub

Public Sub PedPrint_SendTPN()
    
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
    
    Dim blnProgress As Boolean
    
    On Error GoTo PedPrint_SendTPNError
    
    If blnProgress Then ModProgress.StartProgress "TPN naar de apotheek verzenden"
    
    strTo = "c.w.bollen@umcutrecht.nl"
    strFrom = "FunctioneelBeheerMetavision@umcutrecht.nl"
    strSubject = "TPN brief voor " & ModPatient.PatientHospNum & " " & ModPatient.PatientAchterNaam & ", " & ModPatient.PatientVoorNaam
    strHTML = vbNullString
    
    Set objMsg = CreateObject("CDO.Message")
    With objMsg
         
        .To = CStr(strTo)
        .From = CStr(strFrom)
        .Subject = CStr(strSubject)
        .HTMLBody = CStr(strHTML)
        .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'cdoSendUsingPickup=1, cdoSendUsingPort=2, cdoSendUsingExchange=3
        .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.umcutrecht.nl"
        .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        .Configuration.Fields.Update
        
        strFile = Environ("TEMP") & "\TPNbrief_" & ModPatient.PatientHospNum
        strPDF = PrintTPN(False, strFile)
        .AddAttachment strPDF
        strPDFList = strPDFList & strPDF & vbNewLine
        
        .Send
    
    End With
    
    arrPDF = Split(strPDFList, vbNewLine)
    intC = UBound(arrPDF)
    For intN = 0 To intC
        strPDF = Trim(arrPDF(intN))
        If Not strPDF = vbNullString Then Kill strPDF
    Next
    
    Set objMsg = Nothing
    
    If blnProgress Then ModProgress.FinishProgress
    
    ModMessage.ShowMsgBoxInfo "Medicatie is verstuurd naar de apotheek en de patient is opgeslagen"
    
    Exit Sub
    
PedPrint_SendTPNError:
    
    ModLog.LogError "PedPrint_SendTPNError"
    
    On Error Resume Next
    
    ModMessage.ShowMsgBoxError "Medicatie is niet verstuurd naar de apotheek, een foutmelding is verzonden naar functioneel beheer"
    
    arrPDF = Split(strPDFList, vbNewLine)
    intC = UBound(arrPDF)
    For intN = 0 To intC
        strPDF = Trim(arrPDF(intN))
        If Not strPDF = vbNullString Then Kill strPDF
    Next
    
    Set objMsg = Nothing
    
    If blnProgress Then ModProgress.FinishProgress


End Sub

Private Function PrintTPN(ByVal blnPrev As Boolean, ByVal strFile As String) As String

    Dim strPDF As String
    Dim objSheet As Worksheet
    Dim dblGew As Double
    
    dblGew = ModPatient.GetGewichtFromRange

    If dblGew >= CONST_TPN_1 And dblGew <= CONST_TPN_2 Then
        Set objSheet = shtPedPrtTPN2tot6
    Else
        If dblGew <= CONST_TPN_3 Then
            Set objSheet = shtPedPrtTPN7tot15
        Else
            If dblGew <= CONST_TPN_4 Then
                Set objSheet = shtPedPrtTPN16tot30
            Else
                If dblGew <= CONST_TPN_5 Then
                    Set objSheet = shtPedPrtTPN31tot50
                Else
                    Set objSheet = shtPedPrtTPN50
                End If
            End If
        End If
    End If

    objSheet.Unprotect ModConst.CONST_PASSWORD
    
    ModSheet.PrintSheetAllPortrait objSheet
    
    If strFile = vbNullString Then
        PrintSheet objSheet, 1, False, blnPrev
    Else
        strPDF = strFile & ".pdf"
        SaveSheetAsPDF objSheet, strPDF
    End If
    
    PrintTPN = strPDF
    
End Function


