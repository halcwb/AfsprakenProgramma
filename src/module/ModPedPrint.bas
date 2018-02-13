Attribute VB_Name = "ModPedPrint"
Option Explicit

Private Const CONST_TPN_1 As Integer = 2
Private Const CONST_TPN_2 As Integer = 7
Private Const CONST_TPN_3 As Integer = 16
Private Const CONST_TPN_4 As Integer = 31
Private Const CONST_TPN_5 As Integer = 50

Private Const constTPN As String = "_Ped_TPN_Keuze"
Private Const constUserType As String = "_User_Type"

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
    
    Dim strMail As String
    Dim strUser As String
    
    On Error GoTo PedPrint_SendTPNError
    
    strUser = ModRange.GetRangeValue(constUserType, vbNullString)
    If Not (strUser = "Supervisor" Or strUser = "Artsen") And ModSetting.IsProductionDir() Then
        ModMessage.ShowMsgBoxExclam "Er is geen arts ingelogd." & vbNewLine & "Kan de TPN brief niet verzenden!"
        Exit Sub
    End If
    
    strMail = "wkz-algemeen@umcutrecht.nl"
    If Not ModSetting.IsProductionDir() Then strMail = ModMessage.ShowInputBox("Voer een email adres in", vbNullString)
    
    If strMail = vbNullString Then
        ModMessage.ShowMsgBoxExclam "Er moet een email adres worden ingevoerd." & vbNewLine & "Kan de TPN brief niet verzenden!"
        Exit Sub
    End If
    
    ModProgress.StartProgress "TPN brief naar de apotheek verzenden"
    
    strTo = strMail
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
    
    ModProgress.FinishProgress
    
    ModMessage.ShowMsgBoxInfo "TPN brief is verstuurd naar de apotheek."
    
    Exit Sub
    
PedPrint_SendTPNError:
    
    ModLog.LogError "PedPrint_SendTPNError"
    
    On Error Resume Next
    
    ModMessage.ShowMsgBoxError "TPN brief is niet verstuurd naar de apotheek, een foutmelding is verzonden naar functioneel beheer"
    
    arrPDF = Split(strPDFList, vbNewLine)
    intC = UBound(arrPDF)
    For intN = 0 To intC
        strPDF = Trim(arrPDF(intN))
        If Not strPDF = vbNullString Then Kill strPDF
    Next
    
    Set objMsg = Nothing
    
    ModProgress.FinishProgress


End Sub

Private Function PrintTPN(ByVal blnPrev As Boolean, ByVal strFile As String) As String

    Dim strPDF As String
    Dim objSheet As Worksheet
    Dim dblGew As Double
    Dim intTPN As Integer
    
    intTPN = ModRange.GetRangeValue(constTPN, 1)
    dblGew = ModPatient.GetGewichtFromRange()
    
    If intTPN > 1 Then
        If intTPN = 2 Then Set objSheet = shtPedPrtTPN2tot6
        If intTPN = 3 Then Set objSheet = shtPedPrtTPN7tot15
        If intTPN = 4 Then Set objSheet = shtPedPrtTPN16tot30
        If intTPN = 5 Then Set objSheet = shtPedPrtTPN31tot50
        If intTPN = 6 Then Set objSheet = shtPedPrtTPN50
            
    Else
        Set objSheet = shtPedPrtTPN2tot6
        If dblGew >= CONST_TPN_2 Then Set objSheet = shtPedPrtTPN7tot15
        If dblGew >= CONST_TPN_3 Then Set objSheet = shtPedPrtTPN16tot30
        If dblGew >= CONST_TPN_4 Then Set objSheet = shtPedPrtTPN31tot50
        If dblGew > CONST_TPN_5 Then Set objSheet = shtPedPrtTPN50
    
    End If
    
    objSheet.Unprotect ModConst.CONST_PASSWORD
    
    ModSheet.PrintSheetAllPortrait objSheet
    
    If strFile = vbNullString Then
        PrintSheet objSheet, 1, False, blnPrev
    Else
        strPDF = strFile & ".pdf"
        SaveSheetAsPDF objSheet, strPDF, True
    End If
    
    PrintTPN = strPDF
    
End Function


