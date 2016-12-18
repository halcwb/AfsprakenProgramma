Attribute VB_Name = "ModSheets"
Option Explicit

' Get all sheets that act as a User Interface
' Must be visible and protected
Public Function GetUserInterfaceSheets() As Collection
'TODO: Update list of Interface sheets
    Dim col As New Collection
        
    With col
        .Add Item:=shtGuiAcuteOpvang
        .Add Item:=shtGuiInfusen
        .Add Item:=shtGuiIntake
        .Add Item:=shtGuiLab
        .Add Item:=shtGuiMedDisc
        .Add Item:=shtGuiMedicatieIV
        .Add Item:=shtPrtAfspraken
        .Add Item:=shtPrtMedicatie
        .Add Item:=shtPrtTPN16tot30kg
        .Add Item:=shtPrtTPN2tot6kg
        .Add Item:=shtPrtTPN31tot50kg
        .Add Item:=shtPrtTPN7tot15kg
        .Add Item:=shtPrtTPNboven50kg
    
        .Add Item:=shtAanvullendeAfspraken
        .Add Item:=shtAanvullendeAfsprakenPed
        .Add Item:=shtGuiAfspraken
        .Add Item:=shtGuiAfspraken1700
        .Add Item:=shtApotheek
        .Add Item:=shtGuiAcuteOpvangNeo
        .Add Item:=shtGuiLabNeo
        .Add Item:=shtPrint
        .Add Item:=shtGuiWerkBrief
    
    End With
    
    Set GetUserInterfaceSheets = col

End Function

' Get all sheets that do work and are not User Interface
' Must be hidden and not protected
Public Function GetNonInterfaceSheets() As Collection
'TODO: Update list of Calculation sheets

    Dim col As New Collection
    
    With col
        .Add Item:=shtBerConversie
        .Add Item:=shtBerInfusen
        .Add Item:=shtBerIVMed
        .Add Item:=shtBerLab
        .Add Item:=shtBerMedDisc
        .Add Item:=shtBerNormaal
        .Add Item:=shtBerOpm
        .Add Item:=shtBerPO
        .Add Item:=shtBerTemp
        .Add Item:=shtBerTijden
        .Add Item:=shtBerTotalen
        .Add Item:=shtBerTPN
        .Add Item:=shtPatAfsprakenTekst
        .Add Item:=shtPatDetails
        .Add Item:=shtPatData
        .Add Item:=shtTblHeightNL
        .Add Item:=shtTblWeigthNL
        .Add Item:=shtBerTijdenNeo
        .Add Item:=shtTblMedDisc
        .Add Item:=shtTblInfusen
    
        .Add Item:=shtTblMedicatieIV
        .Add Item:=shtTblMedicatieIVNeo
        .Add Item:=shtTblTPOSheet1
        .Add Item:=shtBerekeningen
        .Add Item:=shtBerekeningen1700
        .Add Item:=shtBerLabNeo
        .Add Item:=shtAanvullendeTbl
        .Add Item:=shtAanvullendeBer
        .Add Item:=shtAanvullendeBerPed
        .Add Item:=shtAdvies
        .Add Item:=shtLijsten
        .Add Item:=shtTblTPOSheet1
        .Add Item:=shtTblVoeding
        
        .Add Item:=shtDivPediatrie
        .Add Item:=shtDivNeo
        .Add Item:=shtDivPatient
        .Add Item:=shtWerkBriefActueel
        
        
    End With

    Set GetNonInterfaceSheets = col

End Function

Private Sub HideAndUnProtectNonUserInterfaceSheets()

    Dim col As New Collection
    Dim intCount As Integer
    
    Set col = GetNonInterfaceSheets()
    
    For intCount = 1 To col.Count
        With col(intCount)
            .visible = xlVeryHidden
            .Unprotect PASSWORD:=CONST_PASSWORD
        End With
    Next intCount

    Set col = Nothing

End Sub

Private Sub UnhideNonUserInterfaceSheets()

    Dim col As New Collection
    Dim intCount As Integer
    
    Set col = GetNonInterfaceSheets()
    
    For intCount = 1 To col.Count
        With col(intCount)
            .visible = True
        End With
    Next intCount

    Set col = Nothing

End Sub
