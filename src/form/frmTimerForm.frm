VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTimerForm 
   Caption         =   "Afspraken Timer"
   ClientHeight    =   1533
   ClientLeft      =   42
   ClientTop       =   329
   ClientWidth     =   1834
   OleObjectBlob   =   "frmTimerForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTimerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'————————————–
'Require that all variables be declared
Option Explicit
Public Execute_TimerDrivenMacro As Boolean
'

Private Sub btnResetTimer_Click()
    lblTimer.Caption = ModTimer.TimerCount
End Sub

Private Sub UserForm_Activate()
    'Start_OnTimerMacro
End Sub

Private Sub UserForm_Click()
    MsgBox frmTimerForm.Name
End Sub

Sub Start_OnTimerMacro()
    Execute_TimerDrivenMacro = True
    lblTimer.Caption = ModTimer.TimerCount
    Application.OnTime Time + TimeValue(ModTimer.TimerValue), "OnTimerMacro"
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        MsgBox "De Afspraken Refresh Timer mag niet worden gesloten!", vbExclamation
    End If
End Sub


' ————————————–
