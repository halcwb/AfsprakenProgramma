Attribute VB_Name = "ModSheet"
Option Explicit

Public Function IsUserInterface(shtSheet As Worksheet) As Boolean

    Dim blnGui As Boolean
    
    blnGui = ModString.ContainsCaseSensitive(shtSheet.Name, "Gui")
    blnGui = blnGui Or ModString.ContainsCaseSensitive(shtSheet.Name, "Prt")
    
    IsUserInterface = blnGui

End Function

' Get all sheets that act as a User Interface
' Must be visible and protected
Public Function GetUserInterfaceSheets() As Collection
    Dim col As New Collection
    Dim shtSheet As Worksheet
    
    For Each shtSheet In WbkAfspraken.Sheets
    
        If IsUserInterface(shtSheet) Then
            col.Add shtSheet
        End If
    
    Next shtSheet
    
    Set GetUserInterfaceSheets = col

End Function

' Get all sheets that do work and are not User Interface
' Must be hidden and not protected
Public Function GetNonInterfaceSheets() As Collection
    Dim col As New Collection
    Dim shtSheet As Worksheet
    
    For Each shtSheet In WbkAfspraken.Sheets
    
        If Not IsUserInterface(shtSheet) Then
            col.Add shtSheet
        End If
    
    Next shtSheet
    
    Set GetNonInterfaceSheets = col

End Function

Public Sub UnhideNonUserInterfaceSheets(blnShowProgress As Boolean)

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

Public Sub HideAndUnProtectNonUserInterfaceSheets(blnShowProgress As Boolean)

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
        
        If blnShowProgress Then ModProgress.SetJobPercentage "Verberg Bladen", intC, intN
        intN = intN + 1

    Next shtSheet

    Set colShts = Nothing

End Sub

Public Sub UnprotectUserInterfaceSheets(blnShowProgress As Boolean)

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

Public Sub ProtectUserInterfaceSheets(blnShowProgress As Boolean)
            
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
        
        If blnShowProgress Then ModProgress.SetJobPercentage "Zet Beveiliging", intC, intN
        
    Next objItem

    Set colShts = Nothing

End Sub

Public Sub GoToSheet(shtSheet As Worksheet, strRange As String)

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
Public Sub SelectNeoOrPedSheet(shtPed As Worksheet, shtNeo As Worksheet)

    Dim strPath As String
    Dim strPeli As String
    Dim blnIsDevelop As Boolean
    
    strPath = WbkAfspraken.Path
    strPeli = ModSetting.GetPedDir()
    blnIsDevelop = ModSetting.IsDevelopmentMode()
    
    If ModString.ContainsCaseInsensitive(strPath, strPeli) Or blnIsDevelop Then
        GoToSheet shtPed, "A1"
    Else
        GoToSheet shtNeo, "A1"
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




