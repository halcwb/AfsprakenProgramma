VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPrintAfspraken 
   Caption         =   "Afspraken printen"
   ClientHeight    =   2611
   ClientLeft      =   42
   ClientTop       =   329
   ClientWidth     =   2975
   OleObjectBlob   =   "frmPrintAfspraken.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPrintAfspraken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    frmPrintAfspraken.Hide
End Sub

Private Sub cmdOk_Click()

On Error Resume Next

Application.DisplayAlerts = False
Dim strBed As String, oSheet As Worksheet

strBed = Range("Bednummer").Formula

    frmPrintAfspraken.Hide
    
    Application.Dialogs(xlDialogPrinterSetup).Show
    
    If chkAcuteOpvang.Value Then
        shtGuiAcuteOpvang.Select
        ActiveSheet.PageSetup.CenterHeader = "Bed " & strBed
        Sheets("acuteopvang").PrintOut preview:=False
    End If
    If chkMedicatie.Value Then
        shtPrtMedicatie.Select
        ActiveSheet.PageSetup.LeftHeader = "Bed " & strBed
        Sheets("Medicatie").PrintOut preview:=False
    End If
    If chkTPNBlad.Value Then
         If Val(Range("Gewicht").Text) / 10 < 7 Then
             shtPrtTPN2tot6kg.Select
             ActiveSheet.PageSetup.CenterHeader = "Bed " & strBed
             shtPrtTPN2tot6kg.PrintOut preview:=False
         ElseIf Val(Range("Gewicht").Text) / 10 < 16 Then
             shtPrtTPN7tot15kg.Select
             ActiveSheet.PageSetup.CenterHeader = "Bed " & strBed
             shtPrtTPN7tot15kg.PrintOut preview:=False
          ElseIf Val(Range("Gewicht").Text) / 10 < 31 Then
             shtPrtTPN16tot30kg.Select
             ActiveSheet.PageSetup.CenterHeader = "Bed " & strBed
             shtPrtTPN16tot30kg.PrintOut preview:=False
         ElseIf Val(Range("Gewicht").Text) / 10 <= 50 Then
             shtPrtTPN31tot50kg.Select
             ActiveSheet.PageSetup.CenterHeader = "Bed " & strBed
             shtPrtTPN31tot50kg.PrintOut preview:=False
         ElseIf Val(Range("Gewicht").Text) / 10 > 50 Then
             shtPrtTPNboven50kg.Select
             ActiveSheet.PageSetup.CenterHeader = "Bed " & strBed
             shtPrtTPNboven50kg.PrintOut preview:=False
        End If
    End If
    
    
Application.DisplayAlerts = True

End Sub

