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

Public Function GetInterfaceSheetCount() As Integer

    Dim intN As Integer
    Dim shtSheet As Worksheet
    Dim blnGui As Boolean
    
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
    
    strPath = Application.ActiveWorkbook.Path
    strPeli = ModSetting.GetPedDir()
    blnIsDevelop = ModSetting.IsDevelopmentMode()
    
    If ModString.ContainsCaseInsensitive(strPath, strPeli) Or blnIsDevelop Then
        shtPed.Select
    Else
        shtNeo.Select
    End If
    
End Sub

Public Sub SelectPedOrNeoStartSheet()

    SelectNeoOrPedSheet shtPedGuiMedIV, shtNeoGuiAfspraken
    
End Sub

Public Sub SelectPedOrNeoLabSheet()
    
    SelectNeoOrPedSheet shtPedGuiLab, shtNeoGuiLab
        
End Sub

Public Sub SelectPedOrNeoAfsprExtraSheet()
    
    SelectNeoOrPedSheet shtPedGuiAfsprExta, shtNeoGuiAfsprExtra
        
End Sub




