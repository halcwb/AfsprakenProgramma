VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormCopy1700 
   Caption         =   "17.00 uur Afspraken overnemen naar actuele afspraken"
   ClientHeight    =   14938
   ClientLeft      =   42
   ClientTop       =   378
   ClientWidth     =   17885
   OleObjectBlob   =   "FormCopy1700.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormCopy1700"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    
    Me.Hide

End Sub

Private Sub cmdOk_Click()
    
    ModInfuusbrief.AfsprakenOvernemen Me.optAlles.Value, Me.chkVoeding.Value, Me.chkContinueMedicatie.Value, Me.chkTPN.Value
    Me.Hide

End Sub

Private Sub optAlles_Click()
    
    frmVoeding.Enabled = False
    frmContMed.Enabled = False
    frmTPN.Enabled = False

End Sub

Private Sub optPerBlok_Click()
    
    frmVoeding.Enabled = True
    frmContMed.Enabled = True
    frmTPN.Enabled = True

End Sub

Private Sub SetCaption(ByVal strLbl As String, ByVal strItem As String, intN As Integer, bln1700 As Boolean)
    
    On Error GoTo SetCaptionError
    
    strLbl = IIf(bln1700, Replace(strLbl, "Actueel", "1700"), strLbl)
    strLbl = strLbl & intN
    
    strItem = IIf(bln1700, Replace(strItem, "#", "1700"), Replace(strItem, "#", ""))
    strItem = strItem & intN
    
    Me.Controls(strLbl).Caption = ModRange.GetRangeValue(strItem, vbNullString)

    Exit Sub

SetCaptionError:

    ModMessage.ShowMsgBoxError ModConst.CONST_DEFAULTERROR_MSG
    Application.Cursor = xlDefault
    
    ModLog.LogError "SetCaption " & strLbl & " " & strItem & " " & intN & " " & bln1700
    
    Me.Hide

End Sub

Private Sub UserForm_Activate()

    Dim intN As Integer
    Dim strLbl As String
    Dim strItem As String
    
    strLbl = "lblActueelVoeding"
    strItem = "_NeoVoeding#"
    For intN = 1 To 15
        SetCaption strLbl, strItem, intN, True
        SetCaption strLbl, strItem, intN, False
    Next intN

    strLbl = "lblActueelContMed"
    strItem = "_NeoInfuusContinu#"
    For intN = 1 To 15
        SetCaption strLbl, strItem, intN, True
        SetCaption strLbl, strItem, intN, False
    Next intN

    strLbl = "lblActueelTPN"
    strItem = "_NeoInfuusContinu#"
    For intN = 1 To 12
        SetCaption strLbl, strItem, intN, True
        SetCaption strLbl, strItem, intN, False
    Next intN

End Sub
