VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormPrintAfspraken 
   Caption         =   "Afspraken printen"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2970
   OleObjectBlob   =   "FormPrintAfspraken.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormPrintAfspraken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    Me.Hide

End Sub

Private Sub cmdOK_Click()

    On Error Resume Next

    Application.DisplayAlerts = False
    Dim strBed As String

    strBed = Range("Bednummer").Formula

    Me.Hide
    
    Application.Dialogs(xlDialogPrinterSetup).Show
    
    If chkAcuteOpvang.Value Then
        shtPedGuiAcuut.Select
        shtPedGuiAcuut.PageSetup.CenterHeader = "Bed " & strBed
        shtPedGuiAcuut.PrintOut preview:=False
    End If
    If chkMedicatie.Value Then
        shtPedPrtMedDisc.Select
        shtPedPrtMedDisc.PageSetup.LeftHeader = "Bed " & strBed
        shtPedPrtMedDisc.PrintOut preview:=False
    End If
    If chkTPNBlad.Value Then
        If StringToDouble(Range("Gewicht").Text) / 10 < 7 Then
            shtPedPrtTPN2tot6.Select
            shtPedPrtTPN2tot6.PageSetup.CenterHeader = "Bed " & strBed
            shtPedPrtTPN2tot6.PrintOut preview:=False
        ElseIf StringToDouble(Range("Gewicht").Text) / 10 < 16 Then
            shtPedPrtTPN7tot15.Select
            shtPedPrtTPN7tot15.PageSetup.CenterHeader = "Bed " & strBed
            shtPedPrtTPN7tot15.PrintOut preview:=False
        ElseIf StringToDouble(Range("Gewicht").Text) / 10 < 31 Then
            shtPedPrtTPN16tot30.Select
            shtPedPrtTPN16tot30.PageSetup.CenterHeader = "Bed " & strBed
            shtPedPrtTPN16tot30.PrintOut preview:=False
        ElseIf StringToDouble(Range("Gewicht").Text) / 10 <= 50 Then
            shtPedPrtTPN31tot50.Select
            shtPedPrtTPN31tot50.PageSetup.CenterHeader = "Bed " & strBed
            shtPedPrtTPN31tot50.PrintOut preview:=False
        ElseIf StringToDouble(Range("Gewicht").Text) / 10 > 50 Then
            shtPedPrtTPN50.Select
            shtPedPrtTPN50.PageSetup.CenterHeader = "Bed " & strBed
            shtPedPrtTPN50.PrintOut preview:=False
        End If
    End If
    
    
    Application.DisplayAlerts = True

End Sub

Private Sub CenterForm()

    StartUpPosition = 0
    Left = Application.Left + (0.5 * Application.Width) - (0.5 * Width)
    Top = Application.Top + (0.5 * Application.Height) - (0.5 * Height)

End Sub

Private Sub UserForm_Activate()

    CenterForm

End Sub

