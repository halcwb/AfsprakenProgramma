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
    intC = colShts.count
    
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
    intC = colShts.count
    
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
    intC = colShts.count
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
    intC = colShts.count
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
    
    intN = ActiveWorkbook.Sheets.count
    GetNonInterfaceSheetCount = intN - GetInterfaceSheetCount()

End Function

' Determine the sheet to open with
' If peli or developper then ped sheet
' Else neo sheet
Public Sub SelectNeoOrPedSheet(ByRef shtPed As Worksheet, ByRef shtNeo As Worksheet)

    Dim strPath As String
    Dim strPeli As String
    Dim blnIsDevelop As Boolean
    
    strPath = WbkAfspraken.Path
    strPeli = ModSetting.GetPedDir()
    blnIsDevelop = ModSetting.IsDevelopmentDir()
    
    If ModString.ContainsCaseInsensitive(strPath, strPeli) Or blnIsDevelop Then
        GoToSheet shtPed, "A1"
    Else
        ModNeoInfB.NeoInfB_SelectInfB False
    End If
    
End Sub

Public Sub SelectPedOrNeoStartSheet()

    SelectNeoOrPedSheet shtPedGuiMedIV, shtNeoGuiInfB
    
End Sub

Public Sub SelectPedOrNeoLabSheet()
    
    SelectNeoOrPedSheet shtPedGuiLab, shtNeoGuiLab
        
End Sub

Public Sub SelectPedOrNeoAfsprSheet()
    
    SelectNeoOrPedSheet shtPedGuiAfspr, shtNeoGuiAfspr
        
End Sub




