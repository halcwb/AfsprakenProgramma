Attribute VB_Name = "ModSheet"
Option Explicit

Public Function IsUserInterface(ByRef shtSheet As Worksheet) As Boolean

    Dim blnGui As Boolean
    
    blnGui = ModString.ContainsCaseSensitive(shtSheet.Name, "Gui")
    blnGui = blnGui Or ModString.ContainsCaseSensitive(shtSheet.Name, "Prt")
    
    IsUserInterface = blnGui

End Function

' Get all sheets that act as a User Interface
' Must be visible and protected
Public Function GetUserInterfaceSheets() As Collection
    
    Dim colShts As Collection
    Dim shtSheet As Worksheet
    
    Set colShts = New Collection
    For Each shtSheet In WbkAfspraken.Sheets
    
        If IsUserInterface(shtSheet) Then
            colShts.Add shtSheet
        End If
    
    Next shtSheet
    
    Set GetUserInterfaceSheets = colShts

End Function

' Get all sheets that do work and are not User Interface
' Must be hidden and not protected
Public Function GetNonInterfaceSheets() As Collection
    
    Dim colShts As Collection
    Dim shtSheet As Worksheet
    
    Set colShts = New Collection
    For Each shtSheet In WbkAfspraken.Sheets
    
        If Not IsUserInterface(shtSheet) Then
            colShts.Add shtSheet
        End If
    
    Next shtSheet
    
    Set GetNonInterfaceSheets = colShts

End Function

Public Sub UnhideNonUserInterfaceSheets(ByVal blnShowProgress As Boolean)

    Dim colShts As Collection
    Dim shtSheet As Worksheet
    Dim intN As Integer
    Dim intC As Integer
    
    Set colShts = GetNonInterfaceSheets()
    intN = 1
    intC = colShts.Count
    
    For Each shtSheet In colShts
    
        shtSheet.Visible = True
        
        If blnShowProgress Then ModProgress.SetJobPercentage "Verberg Bladen", intC, intN
        intN = intN + 1

    Next shtSheet

    Set colShts = Nothing

End Sub

Public Sub HideAndUnProtectNonUserInterfaceSheets(ByVal blnShowProgress As Boolean)

    Dim colShts As Collection
    Dim shtSheet As Worksheet
    Dim intN As Integer
    Dim intC As Integer
    
    Set colShts = GetNonInterfaceSheets()
    intN = 1
    intC = colShts.Count
    
    For Each shtSheet In colShts
    
        With shtSheet
            .Visible = xlVeryHidden
            .Unprotect Password:=CONST_PASSWORD
        End With
        
        If blnShowProgress Then ModProgress.SetJobPercentage "Verberg reken bladen ...", intC, intN
        intN = intN + 1

    Next shtSheet

    Set colShts = Nothing

End Sub

Public Sub UnprotectUserInterfaceSheets(ByVal blnShowProgress As Boolean)

    Dim objItem As Worksheet
    Dim colShts As Collection
    Dim intN As Integer
    Dim intC As Integer
    
    Set colShts = GetUserInterfaceSheets()
    intN = 1
    intC = colShts.Count
    For Each objItem In colShts
    
        With objItem
            .Visible = xlSheetVisible
            .EnableSelection = xlNoRestrictions
            .DisplayPageBreaks = True
            .Unprotect ModConst.CONST_PASSWORD
        End With
        
        If blnShowProgress Then ModProgress.SetJobPercentage "Verwijder Beveiliging", intC, intN
        
    Next objItem
            
    Set colShts = Nothing
            
End Sub

Public Sub ProtectUserInterfaceSheets(ByVal blnShowProgress As Boolean)
            
    Dim objItem As Worksheet
    Dim colShts As Collection
    Dim intN As Integer
    Dim intC As Integer
    
    Set colShts = GetUserInterfaceSheets()
    intN = 1
    intC = colShts.Count
    For Each objItem In colShts
    
        With objItem
            .Visible = xlSheetVisible
            .EnableSelection = xlNoSelection
            .DisplayPageBreaks = False
            .Protect ModConst.CONST_PASSWORD
        End With
        
        If blnShowProgress Then ModProgress.SetJobPercentage "Stel beveiliging in", intC, intN
        
    Next objItem

    Set colShts = Nothing

End Sub

Public Sub GoToSheet(ByRef shtSheet As Worksheet, ByVal strRange As String)

    shtSheet.Select
    shtSheet.Range(strRange).Select
    ActiveWindow.ScrollRow = 1

End Sub

Public Function GetInterfaceSheetCount() As Integer

    Dim intN As Integer
    Dim shtSheet As Worksheet
    
    For Each shtSheet In ActiveWorkbook.Sheets
    
        If IsUserInterface(shtSheet) Then
            intN = intN + 1
        End If
    
    Next shtSheet
    
    GetInterfaceSheetCount = intN

End Function

Public Function GetNonInterfaceSheetCount() As Integer

    Dim intN As Integer
    
    intN = ActiveWorkbook.Sheets.Count
    GetNonInterfaceSheetCount = intN - GetInterfaceSheetCount()

End Function

' Determine the sheet to open with
' If peli or developper then ped sheet
' Else neo sheet
Public Sub SelectNeoOrPedSheet(ByRef shtPed As Worksheet, ByRef shtNeo As Worksheet, ByVal blnStartProgress As Boolean)

    Dim blnIsDevelop As Boolean
        
    blnIsDevelop = ModSetting.IsDevelopmentDir()
    
    If ModSetting.IsPed() Or blnIsDevelop Then
        GoToSheet shtPed, "A1"
    Else
        ModNeoInfB.NeoInfB_SelectInfB False, blnStartProgress
    End If
    
End Sub

Public Sub SelectPedOrNeoStartSheet(ByVal blnStartProgress As Boolean)

    SelectNeoOrPedSheet shtPedGuiMedIV, shtNeoGuiInfB, blnStartProgress
    
End Sub

Private Sub Test_SelectPedOrNeoStartSheet()

    SelectPedOrNeoStartSheet True

End Sub


Public Sub SelectPedOrNeoLabSheet()
    
    SelectNeoOrPedSheet shtPedGuiLab, shtNeoGuiLab, False
        
End Sub

Public Sub SelectPedOrNeoAfsprSheet()
    
    SelectNeoOrPedSheet shtPedGuiAfspr, shtNeoGuiAfspr, False
        
End Sub




