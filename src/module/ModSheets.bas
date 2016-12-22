Attribute VB_Name = "ModSheets"
Option Explicit

' Get all sheets that act as a User Interface
' Must be visible and protected
Public Function GetUserInterfaceSheets() As Collection
'TODO: Update list of Interface sheets
    Dim col As New Collection
        
    With col
        ' Ped GUI sheets
        .Add Item:=shtPedGuiAcuut
        .Add Item:=shtPedGuiMedIV
        .Add Item:=shtPedGuiMedDisc
        .Add Item:=shtPedGuiPMenIV
        .Add Item:=shtPedGuiEntTPN
        .Add Item:=shtPedGuiLab
        .Add Item:=shtPedGuiAfsprExta
        ' Ped Print sheets
        .Add Item:=shtPedPrtAfspr
        .Add Item:=shtPedPrtMedDisc
        .Add Item:=shtPedPrtTPN16tot30
        .Add Item:=shtPedPrtTPN2tot6
        .Add Item:=shtPedPrtTPN31tot50
        .Add Item:=shtPedPrtTPN7tot15
        .Add Item:=shtPedPrtTPN50
        ' Neo GUI sheets
        .Add Item:=shtNeoGuiAcuut
        .Add Item:=shtNeoGuiAfspraken
        .Add Item:=shtNeoGuiAfspr1700
        .Add Item:=shtNeoGuiLab
        .Add Item:=shtNeoGuiAfsprExtra
        ' Neo Print sheets
        .Add Item:=shtNeoPrtAfspr
        .Add Item:=shtNeoPrtWerkbr
        .Add Item:=shtNeoPrtWerkbrAct
        .Add Item:=shtNeoPrtApoth
    
    End With
    
    Set GetUserInterfaceSheets = col

End Function

' Get all sheets that do work and are not User Interface
' Must be hidden and not protected
Public Function GetNonInterfaceSheets() As Collection
'TODO: Update list of Calculation sheets

    Dim col As New Collection
    
    With col
        ' Global Berekening sheets
        .Add Item:=shtGlobBerConv
        .Add Item:=shtGlobBerNorm
        .Add Item:=shtGlobBerOpm
        .Add Item:=shtGlobTemp
        ' Pat Data sheets
        .Add Item:=shtPatDetails
        .Add Item:=shtPatData
        .Add Item:=shtPatDataText
        ' Ped Berekening sheets
        .Add Item:=shtPedBerIVenPM
        .Add Item:=shtPedBerMedIV
        .Add Item:=shtPedBerLab
        .Add Item:=shtPedBerMedDisc
        .Add Item:=shtPedBerTot
        .Add Item:=shtPedBerEnt
        .Add Item:=shtPedBerTPN
        .Add Item:=shtPedBerExtraAfspr
        ' Ped Table sheets
        .Add Item:=shtPedTblMedIV
        .Add Item:=shtPedTblMedDisc
        .Add Item:=shtPedTblIV
        .Add Item:=shtPedTblTijden
        .Add Item:=shtPedTblVoed
        .Add Item:=shtPedTblLengte
        .Add Item:=shtPedTblGewicht
        .Add Item:=shtPedTblAfsprExtra
        ' Neo Berekening sheets
        .Add Item:=shtNeoBerAfspr
        .Add Item:=shtNeoBerIV
        .Add Item:=shtNeoBer1700
        .Add Item:=shtNeoBerLab
        .Add Item:=shtNeoBerAdvies
        ' Neo Table sheets
        .Add Item:=shtNeoTblMedIV
        .Add Item:=shtNeoTblTijden
        .Add Item:=shtNeoTblLijst
        .Add Item:=shtNeoTblVoed
        ' Divider sheets
        .Add Item:=shtDivPediatrie
        .Add Item:=shtDivNeo
        .Add Item:=shtDivPatient
        
    End With

    Set GetNonInterfaceSheets = col

End Function

Public Sub HideAndUnProtectNonUserInterfaceSheets()

    Dim col As New Collection
    Dim intCount As Integer
    
    Set col = GetNonInterfaceSheets()
    
    For intCount = 1 To col.Count
        With col(intCount)
            .Visible = xlVeryHidden
            .Unprotect PASSWORD:=CONST_PASSWORD
        End With
    Next intCount

    Set col = Nothing

End Sub

Public Sub UnprotectUserInterfaceSheets()

    Dim objItem As Worksheet
    
    For Each objItem In GetUserInterfaceSheets()
        With objItem
            .EnableSelection = xlNoRestrictions
            .Unprotect ModConst.CONST_PASSWORD
        End With
        
    Next objItem
            
End Sub

Public Sub ProtectUserInterfaceSheets()
            
    Dim objItem As Worksheet
    
    For Each objItem In GetUserInterfaceSheets()
        With objItem
            .EnableSelection = xlNoRestrictions
            .Protect ModConst.CONST_PASSWORD
        End With
        
    Next objItem

End Sub

Public Sub UnhideNonUserInterfaceSheets()

    Dim col As New Collection
    Dim intCount As Integer
    
    Set col = GetNonInterfaceSheets()
    
    For intCount = 1 To col.Count
        With col(intCount)
            .Visible = True
        End With
    Next intCount

    Set col = Nothing

End Sub

Public Sub GoToSheet(shtSheet As Worksheet, strRange As String)

    shtSheet.Select
    shtSheet.Range(strRange).Select
    ActiveWindow.ScrollRow = 1

End Sub

