VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormLog 
   Caption         =   "Kies een gebruikers log"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5775
   OleObjectBlob   =   "FormLog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Logs As Dictionary

Private Sub btnOpenLog_Click()

    Dim strPath  As String

    If lstLogs.ListIndex > -1 Then
        strPath = lstLogs.List(lstLogs.ListIndex)
        strPath = ModSetting.GetLogFileDir & "\" & strPath
        ModLog_OpenLog strPath
    End If

End Sub

Private Sub cboUser_Change()

    Dim strUser As String
    Dim varLog As Variant
    
    lstLogs.Clear
    
    For Each varLog In m_Logs.Keys
        strUser = m_Logs.Item(varLog)
        If strUser = cboUser.Text Then
            lstLogs.AddItem varLog
        End If
    Next

End Sub

Private Sub UserForm_Activate()
    
    Dim strUser As Variant
    
    If m_Logs Is Nothing Then Set m_Logs = ModLog.ModLog_GetLogList()
    
    For Each strUser In m_Logs.Items
        If Not ComboHasItem(cboUser, strUser) Then cboUser.AddItem strUser
    Next

End Sub

Private Function ComboHasItem(objCombo As MSForms.ComboBox, strValue As Variant) As Boolean

    Dim intN As Integer
    Dim intC As Integer
    Dim blnFound As Boolean
    
    blnFound = False
    intC = objCombo.ListCount - 1
    For intN = 0 To intC
        If strValue = objCombo.List(intN) Then
            blnFound = True
            Exit For
        End If
    Next
    
    ComboHasItem = blnFound

End Function

