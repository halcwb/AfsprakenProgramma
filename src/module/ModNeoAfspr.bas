Attribute VB_Name = "ModNeoAfspr"
Option Explicit

Private Const constWondKweek = "_Neo_AfsprD_Wondkweek"

Private Sub EnterText(strCaption As String, strName As String, strRange As String)

    Dim frmInvoer As New FormTekstInvoer
    
    With frmInvoer
        .Caption = strCaption
        .lblNaam.Caption = strName
        .Tekst = ModRange.GetRangeValue(strRange, vbNullString)
        .Show
        If .IsOK Then ModRange.SetRangeValue strRange, .Tekst
    End With
    
    Set frmInvoer = Nothing

End Sub

Public Sub NeoAfspr_WondText()

    EnterText "Voer tekst in ...", "Voer locatie wond(en) in", constWondKweek

End Sub

