VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormColorPicker 
   Caption         =   "Application Colors"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11955
   OleObjectBlob   =   "FormColorPicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const constItem As String = "G"
Private Const constPed As String = "H"
Private Const constNeo As String = "I"

Private Enum Items
    Backgrounds = 2
    Fields = 3
    Labels = 4
    Headers = 5
    Sections = 6
    Totals = 7
    Labs = 8
    Messages = 9
    None = 1
End Enum

Private m_SelectedLabel As MSForms.Label

Private Sub Validate()

    Dim strValid As String
    
    strValid = IIf(cboItem.Value = vbNullString, "Maak een veld selectie", vbNullString)
    strValid = IIf(optNeo.Value = 0 And optPed.Value = 0, "Maak een afdeling selectie", strValid)
    
    cmdOK.Enabled = strValid = vbNullString
    lblValid.Caption = strValid

End Sub

Private Sub ColorRanges()

    If optPed.Value Then
        WriteGroup constPed
        ModColors.ColorPedNeoRanges False
    End If
    
    If optNeo.Value Then
        WriteGroup constNeo
        ModColors.ColorPedNeoRanges True
    End If

End Sub

Private Sub SetLabelBold(ByVal blnBold As Boolean)

    Dim objCtrl As MSForms.Label
    
    For Each objCtrl In frmLabels.Controls
        If objCtrl = m_SelectedLabel Then
            objCtrl.Font.Bold = blnBold
            Exit Sub
        End If
    Next

End Sub

Private Sub SetLabelItalic(ByVal blnItalic As Boolean)

    Dim objCtrl As MSForms.Label
    
    For Each objCtrl In frmLabels.Controls
        If objCtrl = m_SelectedLabel Then
            objCtrl.Font.Italic = blnItalic
            Exit Sub
        End If
    Next

End Sub

Private Function GetItem(ByVal strItem As String) As Items
 
    Dim enmItem As Items

    Select Case strItem
        Case "Backgrounds": enmItem = Items.Backgrounds
        Case "Fields": enmItem = Items.Fields
        Case "Labels": enmItem = Items.Labels
        Case "Headers": enmItem = Items.Headers
        Case "Sections": enmItem = Items.Sections
        Case "Totals": enmItem = Items.Totals
        Case "Labs": enmItem = Items.Labs
        Case "Messages": enmItem = Items.Messages
        Case Else
            enmItem = Items.None
        
    End Select
    
    GetItem = enmItem

End Function

Private Sub SelectLabel(ByRef objLabel As MSForms.Label)

    If Not m_SelectedLabel Is Nothing Then
        m_SelectedLabel.BorderStyle = fmBorderStyleNone
        m_SelectedLabel.SpecialEffect = fmSpecialEffectFlat
        m_SelectedLabel.BorderColor = vbWhite
    End If

    objLabel.BorderStyle = fmBorderStyleSingle
    objLabel.SpecialEffect = fmSpecialEffectRaised
    objLabel.BorderColor = vbYellow
    
    Set m_SelectedLabel = objLabel

End Sub

Private Sub SetLabelColors(ByRef objLabel As MSForms.Label, ByRef objRange As Range)

    objLabel.BackColor = objRange.Interior.Color
    objLabel.ForeColor = objRange.Font.Color
    objLabel.Font.Name = objRange.Font.Name
    objLabel.Font.Size = objRange.Font.Size
    objLabel.Font.Bold = objRange.Font.Bold
    objLabel.Font.Italic = objRange.Font.Italic

End Sub

Private Sub SetRangeColors(ByRef objRange As Range, ByRef objLabel As MSForms.Label)

    objRange.Interior.Color = objLabel.BackColor
    objRange.Font.Color = objLabel.ForeColor
    objRange.Font.Name = objLabel.Font.Name
    objRange.Font.Size = objLabel.Font.Size
    objRange.Font.Bold = objLabel.Font.Bold
    objRange.Font.Italic = objLabel.Font.Italic

End Sub

Private Sub WriteGroup(ByVal strSel As String)

    Dim objRange As Range

    Set objRange = shtGlobSettings.Range(strSel & Items.Backgrounds)
    SetRangeColors objRange, lblBackgrounds
    
    Set objRange = shtGlobSettings.Range(strSel & Items.Fields)
    SetRangeColors objRange, lblFields
    
    Set objRange = shtGlobSettings.Range(strSel & Items.Labels)
    SetRangeColors objRange, lblLabels
    
    Set objRange = shtGlobSettings.Range(strSel & Items.Headers)
    SetRangeColors objRange, lblHeaders
    
    Set objRange = shtGlobSettings.Range(strSel & Items.Sections)
    SetRangeColors objRange, lblSections
    
    Set objRange = shtGlobSettings.Range(strSel & Items.Totals)
    SetRangeColors objRange, lblTotals
    
    Set objRange = shtGlobSettings.Range(strSel & Items.Labs)
    SetRangeColors objRange, lblLabs
    
    Set objRange = shtGlobSettings.Range(strSel & Items.Messages)
    SetRangeColors objRange, lblMessages

End Sub


Private Sub SetGroup(ByVal strSel As String)

    Dim objRange As Range

    Set objRange = shtGlobSettings.Range(strSel & Items.Backgrounds)
    SetLabelColors lblBackgrounds, objRange
    
    Set objRange = shtGlobSettings.Range(strSel & Items.Fields)
    SetLabelColors lblFields, objRange
    
    Set objRange = shtGlobSettings.Range(strSel & Items.Labels)
    SetLabelColors lblLabels, objRange
    
    Set objRange = shtGlobSettings.Range(strSel & Items.Headers)
    SetLabelColors lblHeaders, objRange
    
    Set objRange = shtGlobSettings.Range(strSel & Items.Sections)
    SetLabelColors lblSections, objRange
    
    Set objRange = shtGlobSettings.Range(strSel & Items.Totals)
    SetLabelColors lblTotals, objRange
    
    Set objRange = shtGlobSettings.Range(strSel & Items.Labs)
    SetLabelColors lblLabs, objRange
    
    Set objRange = shtGlobSettings.Range(strSel & Items.Messages)
    SetLabelColors lblMessages, objRange
    
End Sub

Private Sub cboItem_Change()

    Dim objLabel As MSForms.Label
    
    Set objLabel = frmLabels.Controls("lbl" & cboItem.Value)
    
    SelectLabel objLabel
    
    Validate

End Sub

Private Sub cmdApply_Click()

    Me.Hide
    ColorRanges
    Me.Show

End Sub

Private Sub cmdBackGround_Click()

    Dim lngColor As Long
    Dim lngSelected As Long
    
    If Not m_SelectedLabel Is Nothing Then
            lngSelected = m_SelectedLabel.BackColor
            lngColor = ModColors.ShowColorDialog(lngSelected)
            If Not lngColor = -1 Then
                m_SelectedLabel.BackColor = lngColor
            End If
    End If

End Sub

Private Sub cmdCancel_Click()

    Me.Hide

End Sub

Private Sub cmdFont_Click()

    Dim enmItem As Items
    
    If Not m_SelectedLabel Is Nothing Then
        enmItem = GetItem(cboItem.Value)
                
        If optPed.Value Or optNeo.Value Then
            
            FormFontPicker.lblValid.Caption = vbNullString
            FormFontPicker.cboFont.Value = m_SelectedLabel.Font.Name
            FormFontPicker.cboSize.Value = m_SelectedLabel.Font.Size
            FormFontPicker.chkBold.Value = m_SelectedLabel.Font.Bold
            FormFontPicker.chkItalic.Value = m_SelectedLabel.Font.Italic
            
            FormFontPicker.Show
            
            If FormFontPicker.lblValid.Caption = vbNullString Then
                m_SelectedLabel.Font.Name = FormFontPicker.cboFont.Value
                m_SelectedLabel.Font.Size = FormFontPicker.cboSize.Value
                SetLabelBold FormFontPicker.chkBold.Value
                SetLabelItalic FormFontPicker.chkItalic.Value
            End If
            
        End If
    
    End If

End Sub

Private Sub cmdForeGround_Click()

    Dim lngColor As Long
    Dim lngSelected As Long
    
    If Not m_SelectedLabel Is Nothing Then
            lngSelected = m_SelectedLabel.ForeColor
            lngColor = ModColors.ShowColorDialog(lngSelected)
            If Not lngColor = -1 Then
                m_SelectedLabel.ForeColor = lngColor
            End If
    End If

End Sub

Private Sub cmdOK_Click()

    Me.Hide
    ColorRanges

End Sub

Private Sub optPed_Click()

    SetGroup constPed
    Validate

End Sub

Private Sub optNeo_Click()

    SetGroup constNeo
    Validate

End Sub

Private Sub UserForm_Activate()

    Dim objRange As Range

    cboItem.AddItem shtGlobSettings.Range(constItem & Items.Backgrounds).Value2
    cboItem.AddItem shtGlobSettings.Range(constItem & Items.Fields).Value2
    cboItem.AddItem shtGlobSettings.Range(constItem & Items.Labels).Value2
    cboItem.AddItem shtGlobSettings.Range(constItem & Items.Headers).Value2
    cboItem.AddItem shtGlobSettings.Range(constItem & Items.Sections).Value2
    cboItem.AddItem shtGlobSettings.Range(constItem & Items.Totals).Value2
    cboItem.AddItem shtGlobSettings.Range(constItem & Items.Labs).Value2
    cboItem.AddItem shtGlobSettings.Range(constItem & Items.Messages).Value2
    
    SetGroup constItem
    
    Validate

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    lblValid.Caption = "Cancel"
    Cancel = True
    Me.Hide

End Sub
