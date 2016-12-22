Attribute VB_Name = "ModTimer"
Option Explicit

Public Const TimerValue As String = "00:00:05"
Public Const TimerCount As Integer = 10

Public Sub OnTimerMacro()

    Dim frmTimerForm As New FormTimerForm
    
    frmTimerForm.lblTimer.Caption = CInt(frmTimerForm.lblTimer.Caption) - 1
    'frmTimerForm.Repaint
    
    If frmTimerForm.lblTimer.Caption = "0" Then
        Dim strMsg As String
        Dim strBed As String
        Dim strFileName As String
        Dim dtmVersion As Date
        
        If Not Range("BedNummer").Value = 0 Then
            strBed = Range("Bednummer").Formula
            strFileName = GetPatientDataFile(strBed)
            
            dtmVersion = FileSystem.FileDateTime(strFileName)
            If Not dtmVersion = Range("AfsprakenVersie").Value Then
                strMsg = strMsg & "De afspraken zijn inmiddels gewijzigd!" & vbNewLine
                strMsg = strMsg & "Open het bed opnieuw en voer de afspraken" & vbNewLine
                strMsg = strMsg & "opnieuw in"
                MsgBox strMsg, vbCritical
                
                OpenPatientLijst
            End If
        End If
        
        ' At the end restart timer
        frmTimerForm.lblTimer.Caption = ModTimer.TimerCount
        Application.OnTime Time + TimeValue(ModTimer.TimerValue), "OnTimerMacro"
    Else
        Application.OnTime Time + TimeValue(ModTimer.TimerValue), "OnTimerMacro"
    
    End If
    
    Set frmTimerForm = Nothing
    
End Sub

