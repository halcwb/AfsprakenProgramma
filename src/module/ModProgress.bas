Attribute VB_Name = "ModProgress"
Option Explicit

Public Sub SetJobPercentage(ByVal strJob As String, ByVal lngTot As Long, ByVal lngProg As Long)
    
    Dim intPerc As Integer
    Dim dblPerc As Double
    
    On Error Resume Next
    
    dblPerc = (CDbl(lngProg) / CDbl(lngTot)) * 100
    If dblPerc > 100 Then
        intPerc = 100
    ElseIf dblPerc < 0 Then
        intPerc = 0
    Else
        intPerc = Int((CDbl(lngProg) / CDbl(lngTot)) * 100)
    End If
    
    If intPerc <= 100 Then FormProgress.SetJobPercentage strJob, intPerc

End Sub

Public Sub StartProgress(ByVal strTitle As String)

    FormProgress.SetCaption strTitle
    FormProgress.SetJobPercentage vbNullString, 0
    FormProgress.Show vbModeless

End Sub

Public Sub FinishProgress()

    FormProgress.Hide

End Sub
