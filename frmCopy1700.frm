VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCopy1700 
   Caption         =   "17.00 uur Afspraken overnemen naar actuele afspraken"
   ClientHeight    =   14938
   ClientLeft      =   42
   ClientTop       =   378
   ClientWidth     =   17885
   OleObjectBlob   =   "frmCopy1700.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCopy1700"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOk_Click()
    mAfspraken1700.AfsprakenOvernemen Me.optAlles.Value, Me.chkVoeding.Value, Me.chkContinueMedicatie.Value, Me.chkTPN.Value
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

Sub UserForm_Activate()
    Me.lblActueelVoeding1.Caption = Range("_NeoVoeding1").Value
    Me.lblActueelVoeding2.Caption = Range("_NeoVoeding2").Value
    Me.lblActueelVoeding3.Caption = Range("_NeoVoeding3").Value
    Me.lblActueelVoeding4.Caption = Range("_NeoVoeding4").Value
    Me.lblActueelVoeding5.Caption = Range("_NeoVoeding5").Value
    Me.lblActueelVoeding6.Caption = Range("_NeoVoeding6").Value
    Me.lblActueelVoeding7.Caption = Range("_NeoVoeding7").Value
    Me.lblActueelVoeding8.Caption = Range("_NeoVoeding8").Value
    Me.lblActueelVoeding9.Caption = Range("_NeoVoeding9").Value
    Me.lblActueelVoeding10.Caption = Range("_NeoVoeding10").Value
    Me.lblActueelVoeding11.Caption = Range("_NeoVoeding11").Value
    Me.lblActueelVoeding12.Caption = Range("_NeoVoeding12").Value
    Me.lblActueelVoeding13.Caption = Range("_NeoVoeding13").Value
    Me.lblActueelVoeding14.Caption = Range("_NeoVoeding14").Value
    Me.lblActueelVoeding15.Caption = Range("_NeoVoeding15").Value
    
    Me.lbl1700Voeding1.Caption = Range("_NeoVoeding1700_1").Value
    Me.lbl1700Voeding2.Caption = Range("_NeoVoeding1700_2").Value
    Me.lbl1700Voeding3.Caption = Range("_NeoVoeding1700_3").Value
    Me.lbl1700Voeding4.Caption = Range("_NeoVoeding1700_4").Value
    Me.lbl1700Voeding5.Caption = Range("_NeoVoeding1700_5").Value
    Me.lbl1700Voeding6.Caption = Range("_NeoVoeding1700_6").Value
    Me.lbl1700Voeding7.Caption = Range("_NeoVoeding1700_7").Value
    Me.lbl1700Voeding8.Caption = Range("_NeoVoeding1700_8").Value
    Me.lbl1700Voeding9.Caption = Range("_NeoVoeding1700_9").Value
    Me.lbl1700Voeding10.Caption = Range("_NeoVoeding1700_10").Value
    Me.lbl1700Voeding11.Caption = Range("_NeoVoeding1700_11").Value
    Me.lbl1700Voeding12.Caption = Range("_NeoVoeding1700_12").Value
    Me.lbl1700Voeding13.Caption = Range("_NeoVoeding1700_13").Value
    Me.lbl1700Voeding14.Caption = Range("_NeoVoeding1700_14").Value
    Me.lbl1700Voeding15.Caption = Range("_NeoVoeding1700_15").Value
    
    Me.lblActueelContMed1.Caption = Range("_NeoInfuusContinu1").Value
    Me.lblActueelContMed2.Caption = Range("_NeoInfuusContinu2").Value
    Me.lblActueelContMed3.Caption = Range("_NeoInfuusContinu3").Value
    Me.lblActueelContMed4.Caption = Range("_NeoInfuusContinu4").Value
    Me.lblActueelContMed5.Caption = Range("_NeoInfuusContinu5").Value
    Me.lblActueelContMed6.Caption = Range("_NeoInfuusContinu6").Value
    Me.lblActueelContMed7.Caption = Range("_NeoInfuusContinu7").Value
    Me.lblActueelContMed8.Caption = Range("_NeoInfuusContinu8").Value
    Me.lblActueelContMed9.Caption = Range("_NeoInfuusContinu9").Value
    Me.lblActueelContMed10.Caption = Range("_NeoInfuusContinu27").Value
    Me.lblActueelContMed11.Caption = Range("_NeoInfuusContinu10").Value
    Me.lblActueelContMed12.Caption = Range("_NeoInfuusContinu11").Value
    Me.lblActueelContMed13.Caption = Range("_NeoInfuusContinu12").Value
    Me.lblActueelContMed14.Caption = Range("_NeoInfuusContinu13").Value
    Me.lblActueelContMed15.Caption = Range("_NeoInfuusContinu14").Value
    
    Me.lbl1700ContMed1.Caption = Range("_NeoInfuusContinu1700_1").Value
    Me.lbl1700ContMed2.Caption = Range("_NeoInfuusContinu1700_2").Value
    Me.lbl1700ContMed3.Caption = Range("_NeoInfuusContinu1700_3").Value
    Me.lbl1700ContMed4.Caption = Range("_NeoInfuusContinu1700_4").Value
    Me.lbl1700ContMed5.Caption = Range("_NeoInfuusContinu1700_5").Value
    Me.lbl1700ContMed6.Caption = Range("_NeoInfuusContinu1700_6").Value
    Me.lbl1700ContMed7.Caption = Range("_NeoInfuusContinu1700_7").Value
    Me.lbl1700ContMed8.Caption = Range("_NeoInfuusContinu1700_8").Value
    Me.lbl1700ContMed9.Caption = Range("_NeoInfuusContinu1700_9").Value
    Me.lbl1700ContMed10.Caption = Range("_NeoInfuusContinu1700_27").Value
    Me.lbl1700ContMed11.Caption = Range("_NeoInfuusContinu1700_10").Value
    Me.lbl1700ContMed12.Caption = Range("_NeoInfuusContinu1700_11").Value
    Me.lbl1700ContMed13.Caption = Range("_NeoInfuusContinu1700_12").Value
    Me.lbl1700ContMed14.Caption = Range("_NeoInfuusContinu1700_13").Value
    Me.lbl1700ContMed15.Caption = Range("_NeoInfuusContinu1700_14").Value
    
    Me.lblActueelTPN1.Caption = Range("_NeoInfuusContinu15").Value
    Me.lblActueelTPN2.Caption = Range("_NeoInfuusContinu16").Value
    Me.lblActueelTPN3.Caption = Range("_NeoInfuusContinu17").Value
    Me.lblActueelTPN4.Caption = Range("_NeoInfuusContinu18").Value
    Me.lblActueelTPN5.Caption = Range("_NeoInfuusContinu19").Value
    Me.lblActueelTPN6.Caption = Range("_NeoInfuusContinu20").Value
    Me.lblActueelTPN7.Caption = Range("_NeoInfuusContinu21").Value
    Me.lblActueelTPN8.Caption = Range("_NeoInfuusContinu22").Value
    Me.lblActueelTPN9.Caption = Range("_NeoInfuusContinu23").Value
    Me.lblActueelTPN10.Caption = Range("_NeoInfuusContinu24").Value
    Me.lblActueelTPN11.Caption = Range("_NeoInfuusContinu25").Value
    Me.lblActueelTPN12.Caption = Range("_NeoInfuusContinu26").Value
    
    Me.lbl1700TPN1.Caption = Range("_NeoInfuusContinu1700_15").Value
    Me.lbl1700TPN2.Caption = Range("_NeoInfuusContinu1700_16").Value
    Me.lbl1700TPN3.Caption = Range("_NeoInfuusContinu1700_17").Value
    Me.lbl1700TPN4.Caption = Range("_NeoInfuusContinu1700_18").Value
    Me.lbl1700TPN5.Caption = Range("_NeoInfuusContinu1700_19").Value
    Me.lbl1700TPN6.Caption = Range("_NeoInfuusContinu1700_20").Value
    Me.lbl1700TPN7.Caption = Range("_NeoInfuusContinu1700_21").Value
    Me.lbl1700TPN8.Caption = Range("_NeoInfuusContinu1700_22").Value
    Me.lbl1700TPN9.Caption = Range("_NeoInfuusContinu1700_23").Value
    Me.lbl1700TPN10.Caption = Range("_NeoInfuusContinu1700_24").Value
    Me.lbl1700TPN11.Caption = Range("_NeoInfuusContinu1700_25").Value
    Me.lbl1700TPN12.Caption = Range("_NeoInfuusContinu1700_26").Value
End Sub
