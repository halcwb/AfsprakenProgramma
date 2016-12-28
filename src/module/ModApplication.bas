Attribute VB_Name = "ModApplication"
Option Explicit

Private blnDontClose As Boolean

Public Enum EnumAppLanguage
    Dutch = 1043
    English = 0
End Enum

Public Sub SetDontClose(blnClose As Boolean)
    
    blnDontClose = blnClose

End Sub

Public Sub SetToDevelopmentMode()

    Dim objSheet As Worksheet
    
    blnDontClose = True
        
    ModSheet.UnprotectUserInterfaceSheets
    ModSheet.UnhideNonUserInterfaceSheets
    
    Application.DisplayFormulaBar = True
    
    For Each objSheet In ModSheet.GetNonInterfaceSheets()
        objSheet.Activate
        SetWindowToCloseApp Windows(1)
    Next
    
    For Each objSheet In ModSheet.GetUserInterfaceSheets()
        objSheet.Activate
        SetWindowToCloseApp Windows(1)
    Next
    
    blnDontClose = False
    ModSetting.SetDevelopmentMode True
    
    Application.Cursor = xlDefault

End Sub

Public Sub CloseAfspraken()

    Dim strAction As String, strParams() As Variant
    strAction = "ModApplication.CloseAfspraken"
    strParams = Array()
    
    ModLog.LogActionStart strAction, strParams
    
    Dim objWindow As Window

    Application.Cursor = xlWait
    Application.DisplayAlerts = False
    
    For Each objWindow In Application.Windows
        SetWindowToCloseApp objWindow
    Next
 
    Toolbars("Afspraken").Visible = False
    
    With Application
         .Caption = vbNullString
         .DisplayFormulaBar = True
         .Cursor = xlDefault
         If Not blnDontClose Then .Quit
    End With
        
    ModLog.LogActionEnd strAction
            
End Sub

Private Sub SetWindow(objWindow As Window, blnDisplay As Boolean)

    blnDisplay = blnDisplay Or ModSetting.IsDevelopmentMode()

    With objWindow
        .DisplayWorkbookTabs = blnDisplay
        .DisplayGridlines = blnDisplay
        .DisplayHeadings = blnDisplay
        .DisplayOutline = blnDisplay
        .DisplayZeros = blnDisplay
    End With

End Sub

Public Sub SetWindowToCloseApp(objWindow As Window)
    
    SetWindow objWindow, True

End Sub

Public Sub SetWindowToOpenApp(objWindow As Window)
    
    SetWindow objWindow, False

End Sub

Public Sub InitializeAfspraken()

    Dim strError As String
    Dim blnLog As Boolean
    Dim strAction As String
    Dim strParams() As Variant
    Dim objWindow As Window
    
    On Error GoTo InitializeError
    
    strAction = "ModApplication.InitializeAfspraken"
    strParams = Array()
    
    ModLog.LogActionStart strAction, strParams
    
    Application.Cursor = xlWait
    WbkAfspraken.Activate

    ModSheet.ProtectUserInterfaceSheets
    ModSheet.HideAndUnProtectNonUserInterfaceSheets

    SetCaptionAndHideBars
    ActiveWindow.DisplayWorkbookTabs = ModSetting.IsDevelopmentMode()

    For Each objWindow In Application.Windows
        SetWindowToOpenApp objWindow
    Next
    
'   Zorg ervoor dat niet per ongeluk een lege patient naar een bed wordt weggeschreven
    ModRange.SetRangeValue "BedNummer", 0
    ModRange.SetRangeValue "AfsprakenVersie", vbNullString
    
    SetDateToDayFormula
    
    ModPatient.ClearPatient False
    
    ModSetting.SetDevelopmentMode False
    
    Application.Cursor = xlDefault
    
    ModLog.LogActionEnd strAction
    
    Exit Sub
    
InitializeError:
    
    Application.Cursor = xlDefault

    strError = ModConst.CONST_DEFAULTERROR_MSG & vbNewLine & " Kan de applicatie niet opstarten"
    ModMessage.ShowMsgBoxError strError
    
    blnLog = ModSetting.GetEnableLogging
    ModLog.EnableLogging
    strError = strError & vbNewLine & strAction
    ModLog.LogToFile ModSetting.GetLogPath(), Error, strError
    If Not blnLog Then ModLog.DisableLogging
    
End Sub

Public Sub SetDateToDayFormula()

    ModRange.SetRangeFormula "AfspraakDatum", GetToDayFormula()

End Sub

Private Sub SetCaptionAndHideBars()

    Dim blnIsDevelop As Boolean
    
    blnIsDevelop = ModSetting.IsDevelopmentMode()
    
    SetApplicationTitle
    
    With Application
         .DisplayFormulaBar = blnIsDevelop
         .DisplayStatusBar = blnIsDevelop
         .DisplayScrollBars = True
         .DisplayFormulaBar = blnIsDevelop
    End With
    
End Sub

Public Sub SetApplicationTitle()

    Dim strTitle As String
    Dim strBed As String
    Dim strVn As String
    Dim strAn As String
    
    strTitle = ModConst.CONST_APPLICATION_NAME
    strBed = ModRange.GetRangeValue("_bed", "")
    strVn = ModRange.GetRangeValue("_VoorNaam", "")
    strAn = ModRange.GetRangeValue("_AchterNaam", "")
    
    If Not strBed = "0" Then
        strTitle = strTitle & " Patient: " & strAn & " " & strVn & ", Bed: " & strBed
    End If
    
    Application.Caption = strTitle

End Sub

Public Function GetLanguage() As EnumAppLanguage
    
    Dim enmLanguage As EnumAppLanguage
    
    Select Case Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    Case EnumAppLanguage.Dutch: enmLanguage = Dutch
    Case Else: enmLanguage = EnumAppLanguage.English
    End Select
    
    GetLanguage = enmLanguage

End Function

Private Function GetToDayFormula() As String
    
    Dim strToDay As String

    Select Case GetLanguage()
    Case EnumAppLanguage.Dutch: strToDay = "= Vandaag()"
    Case Else: strToDay = "= ToDay()"
    End Select
    
    GetToDayFormula = strToDay

End Function

Private Sub Test()

    MsgBox GetLanguage()

End Sub


