VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormPatient 
   Caption         =   "Nieuwe patient"
   ClientHeight    =   3318
   ClientLeft      =   42
   ClientTop       =   329
   ClientWidth     =   7308
   OleObjectBlob   =   "FormPatient.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const c_MinLengte = 25

Public curDate As Date, aDate As Date, opnDate As Date, bDate As Date
'Private objPatientNummer As HL7nawXControl.HL7nawX

Private Sub cmdCancel_Click()

    Me.Hide

End Sub

Private Sub cmdOk_Click()

    If txtOpnDat.Value = vbNullString Or txtANaam.Value = vbNullString Or txtGebDat.Value = vbNullString _
       Or txtGew.Value = vbNullString Or txtLengte.Value = vbNullString Then
        MsgBox prompt:="De opname datum, Achternaam, gewicht en lengte moeten in elk " & _
                        "geval worden ingevoerd", Buttons:=vbOKOnly, Title:="Informedica"
        txtOpnDat.SetFocus
    Else
        Range("Opndatum").Formula = DateValue(txtOpnDat.Value)
        Range("AfspraakDatum").Value = curDate
        Range("PatNummer").Value = txtPatNum.Value
        Range("_AchterNaam").Value = txtANaam.Value
        Range("_VoorNaam").Value = txtVNaam.Text
        If IsDate(txtGebDat.Value) Then
            Range("GebDatum").Formula = DateValue(txtGebDat.Value)
        Else
            Range("GebDatum").Formula = (txtGebDat.Value)
        End If
        Range("_Weken").Value = txtWeken.Value
        Range("_Dagen").Value = txtDagen.Value
        Range("Gewicht").Value = txtGew.Value * 10
        Range("_Gewicht").Value = txtGew.Value * 1
        Range("Lengte").Value = txtLengte.Value
        'Beademing sheets are not in workbook anymore
        '        If optCMV.Value Then
        '            Worksheets("BeademingCV").Visible = xlSheetVisible
        '            Worksheets("BeademingCV").Activate
        '            Worksheets("BeademingHFO").Visible = xlSheetVeryHidden
        '            Worksheets("BeademingPLV").Visible = xlSheetVeryHidden
        '            Range("Beademing").Value = "CMV"
        '        Else
        '            If optHFOV.Value Then
        '                Worksheets("BeademingCV").Visible = xlSheetVeryHidden
        '                Worksheets("BeademingHFO").Visible = xlSheetVisible
        '                Worksheets("BeademingHFO").Activate
        '                Worksheets("BeademingPLV").Visible = xlSheetVeryHidden
        '                Range("Beademing").Value = "HFOV"
        '            Else
        '                Worksheets("BeademingCV").Visible = xlSheetVeryHidden
        '                Worksheets("BeademingHFO").Visible = xlSheetVeryHidden
        '                Worksheets("BeademingPLV").Visible = xlSheetVisible
        '                Worksheets("BeademingPLV").Activate
        '                Range("Beademing").Value = "PLV"
        '            End If
        '        End If

        SelectTPN
        
    End If

    Me.Hide

End Sub

Private Sub txtGebDat_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)
    On Error GoTo Hell
    
    If Not IsDate(txtGebDat.Value) Then
        MsgBox prompt:="Dit is geen geldige datum", Buttons:=vbCritical, _
               Title:="Informedica"
        ''Cancel=true
        txtGebDat.SetFocus
        txtGebDat.Value = vbNullString
    
    Else
        
        If Not txtOpnDat.Value = vbNullString And _
           DateValue(txtOpnDat.Value) < DateValue(txtGebDat.Value) Then
            MsgBox prompt:="De opname datum kan niet eerder zijn" & _
                            " dan de geboortedatum", Buttons:=vbCritical, _
                   Title:="Informedica"
            'Cancel=true
            txtGebDat.SetFocus
            txtGebDat.Value = vbNullString
        
        Else
            
            If DateValue(txtGebDat.Value) > curDate Then
                MsgBox prompt:="De geboorte datum kan niet later zijn" & _
                                " dan de huidige datum", Buttons:=vbCritical, _
                       Title:="Informedica"
                'Cancel=true
                txtGebDat.SetFocus
                txtGebDat.Value = vbNullString
            End If
        End If
    End If
    
Hell:
End Sub

Private Sub txtOpnDat_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    On Error GoTo Hell
    curDate = DateTime.Date
        
    If Not IsDate(txtOpnDat.Value) Then
        MsgBox prompt:="Dit is geen geldige datum", Buttons:=vbCritical, _
               Title:="Informedica"
        'Cancel=true
        txtOpnDat.SetFocus
        txtOpnDat.Value = vbNullString
    Else
        If DateValue(txtOpnDat.Value) < DateValue(txtGebDat.Value) Then
            MsgBox prompt:="De opname datum kan niet eerder zijn" & _
                            " dan de geboortedatum", Buttons:=vbCritical, _
                   Title:="Informedica"
            'Cancel=true
            txtOpnDat.SetFocus
            txtOpnDat.Value = vbNullString
        Else
            If DateValue(txtOpnDat.Value) > curDate Then
                MsgBox prompt:="De opname datum kan niet later zijn " & _
                                "dan de huidige datum", Buttons:=vbCritical, _
                       Title:="Informedica"
                'Cancel=true
                txtOpnDat.SetFocus
                txtOpnDat.Value = vbNullString
            End If
        End If
    End If

Hell:
End Sub
'
Private Sub Toevoegen()

    Dim intI As Integer, intJ As Integer, intEmpty As Integer
    intEmpty = 0
    
    On Error GoTo Hell
    
    With Sheets("Patienten")
        For intJ = 4 To .Range("a1").CurrentRegion.Columns.Count
            If txtANaam.Text = .Cells(2, intJ).Value Then
                intEmpty = intJ
            End If
        Next intJ
        If intEmpty = 0 Then
            intEmpty = .Range("a1").CurrentRegion.Columns.Count + 1
        End If
        For intI = 2 To .Range("a1").CurrentRegion.Rows.Count
            .Cells(intI, intEmpty).Formula = Range(.Cells(intI, 1).Value).Value
        Next intI
    End With
    
Hell:
    Resume Next
End Sub

Private Sub UserForm_Activate()
    UserForm_Initialize
End Sub

Private Sub UserForm_Initialize()

    txtPatNum.Text = Range("PatNummer").Value
    
    curDate = DateTime.Date
    txtOpnDat.Value = curDate
    txtANaam.Value = Range("_AchterNaam").Value
    txtVNaam.Value = Range("_VoorNaam").Value
    txtGebDat.Value = CDate(Range("GebDatum").Value)
    txtGew.Value = Range("Gewicht").Value / 10
    txtLengte.Value = Range("Lengte").Value
    txtWeken.Value = Range("_Weken").Value
    txtDagen.Value = Range("_Dagen").Value

End Sub

Private Sub txtLengte_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)
    If Not IsNumeric(txtLengte.Value) Or txtLengte.Value <= (c_MinLengte / 100) Or _
       (txtLengte.Value > 2 And txtLengte.Value < c_MinLengte) Or txtLengte.Value > 200 Then
        If Not txtLengte.Value = vbNullString Then
            MsgBox prompt:="Dit is geen geldig lengte", Buttons:=vbCritical, _
                   Title:="Informedica"
            'Cancel = True
            txtLengte.SetFocus
            txtLengte.Value = vbNullString
        End If
    Else
        If txtLengte.Value < c_MinLengte Then
            txtLengte.Value = txtLengte.Value * 100
        End If
    End If

End Sub

Private Sub txtGew_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    If Not IsNumeric(txtGew.Value) Or txtGew.Value <= 0 Or _
       (txtGew.Value > 100 And txtGew.Value < 1500) Then
        If Not txtGew.Text = vbNullString Then
            MsgBox prompt:="Dit is geen geldig gewicht", Buttons:=vbCritical, _
                   Title:="Informedica"
            'Cancel = True
            txtGew.SetFocus
            txtGew.Value = vbNullString
        End If
    Else
        If txtGew.Value > 500 Then
            txtGew.Value = txtGew.Value / 1000
        End If
    End If
            
End Sub

Private Sub txtGew_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.CorrectNumberAscii(intKey)
    
End Sub

