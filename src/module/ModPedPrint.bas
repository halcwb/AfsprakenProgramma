Attribute VB_Name = "ModPedPrint"
Option Explicit

Public Sub SaveAndPrintAfspraken()

    Dim frmPrintAfspraken As New FormPrintAfspraken
    
    ModBed.CloseBed (True)
    frmPrintAfspraken.Show
    
    Set frmPrintAfspraken = Nothing
    
End Sub

Public Sub AfsprakenPrinten()

    shtNeoPrtAfspr.PrintPreview

End Sub

Public Sub WerkBriefPrinten()
        
    With shtNeoPrtWerkbr
        .Unprotect ModConst.CONST_PASSWORD
        .PrintPreview
        .Protect ModConst.CONST_PASSWORD
    End With

End Sub
