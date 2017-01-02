Attribute VB_Name = "ModProgress"
Option Explicit

Public Sub SetJobPercentage(strJob As String, intTot As Integer, intProg As Integer)
    
    Dim intPerc As Integer
    
    intPerc = Int((CDbl(intProg) / CDbl(intTot)) * 100)

    If intPerc <= 100 Then FormProgress.SetJobPercentage strJob, intPerc

End Sub

Public Sub StartProgress(strTitle As String)

    FormProgress.SetCaption strTitle
    FormProgress.SetJobPercentage "", 0
    FormProgress.Show vbModeless

End Sub

Public Sub FinishProgress()

    FormProgress.Hide

End Sub
