Attribute VB_Name = "ModPatienten"
Option Explicit

Private colPiPatienten As Collection
Private intCount As Integer

Public Sub PaPatientenLijst()
    
    Dim strIndex As String
    
    Set colPiPatienten = New Collection
    
    Set colPiPatienten = oPuPatientenCollectie
    
    With frmPatLijst
        Application.Cursor = xlWait
        .lstPatienten.Clear
        For intCount = 1 To colPiPatienten.Count
            .lstPatienten.AddItem colPiPatienten(intCount)
        Next intCount
        Application.Cursor = xlDefault
        .Show
        If .lstPatienten.ListIndex > -1 Then
            Application.Cursor = xlWait
            strIndex = VBA.Left$(.lstPatienten.Text, CONST_BEDNAME_LENGTH)
            Call BeOpenenBed(strIndex)
            Application.Cursor = xlDefault
        End If
        .lstPatienten.Clear
    End With
    
    Set colPiPatienten = Nothing
    
    OpenStartSheet
    
End Sub
