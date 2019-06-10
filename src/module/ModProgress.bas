Attribute VB_Name = "ModProgress"
Option Explicit

Public Sub SetJobPercentage(ByVal strJob As String, ByVal intTot As Long, ByVal intProg As Long)
    
    Dim intPerc As Integer
    
    intPerc = Int((CDbl(intProg) / CDbl(intTot)) * 100)

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
