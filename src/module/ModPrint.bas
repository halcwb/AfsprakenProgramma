Attribute VB_Name = "ModPrint"
Option Explicit

Public Sub SaveAndPrintAfspraken()

    Dim frmPrintAfspraken As New FormPrintAfspraken
    
    ModBed.SluitBed
    frmPrintAfspraken.Show
    
    Set frmPrintAfspraken = Nothing
    
End Sub

Public Sub PrintLabAanvragen()

    With Application
        .DisplayAlerts = False
        'TODO: Link controleren op werking
        .Workbooks.Open "G:\Zorgeenh\Pelikaan\ICAP Data\LabAanvragen.xls", True, True
        .ActiveWorkbook.Sheets("Unit 1").PrintOut
        .ActiveWorkbook.Sheets("Unit 2").PrintOut
        .Workbooks("LabAanvragen.xls").Close
    End With

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
