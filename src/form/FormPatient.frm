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

Private Sub cmdCancel_Click()

    Me.Hide

End Sub

Private Sub cmdOk_Click()

    Dim dtmDate As Date
    
    dtmDate = Now()

    If txtOpnDat.Value = vbNullString Or txtANaam.Value = vbNullString Or txtGebDat.Value = vbNullString _
       Or txtGew.Value = vbNullString Or txtLengte.Value = vbNullString Then
        MsgBox prompt:="De opname datum, Achternaam, gewicht en lengte moeten in elk " & _
                        "geval worden ingevoerd", Buttons:=vbOKOnly, Title:="Informedica"
        txtOpnDat.SetFocus
    Else
        ModRange.SetRangeValue "Opndatum", DateValue(txtOpnDat.Value)
        ModRange.SetRangeValue "AfspraakDatum", dtmDate
        ModRange.SetRangeValue ModConst.CONST_RANGE_PATNUM, txtPatNum.Value
        ModRange.SetRangeValue ModConst.CONST_RANGE_AN, txtANaam.Value
        ModRange.SetRangeValue ModConst.CONST_RANGE_VN, txtVNaam.Text
        If IsDate(txtGebDat.Value) Then
            ModRange.SetRangeValue ModConst.CONST_RANGE_GEBDAT, DateValue(txtGebDat.Value)
        Else
            ModRange.SetRangeValue ModConst.CONST_RANGE_GEBDAT, (txtGebDat.Value)
        End If
        ModRange.SetRangeValue "_Weken", txtWeken.Value
        ModRange.SetRangeValue "_Dagen", txtDagen.Value
        ModRange.SetRangeValue "Gewicht", txtGew.Value * 10
        ModRange.SetRangeValue "_Gewicht", txtGew.Value * 1
        ModRange.SetRangeValue "Lengte", txtLengte.Value

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
            
            If DateValue(txtGebDat.Value) > Now() Then
                MsgBox prompt:="De geboorte datum kan niet later zijn" & _
                                " dan de huidige datum", Buttons:=vbCritical, _
                       Title:="Informedica"
                'Cancel=true
                txtGebDat.SetFocus
                txtGebDat.Value = vbNullString
            End If
        End If
    End If
    
    Exit Sub
    
Hell:
End Sub

Private Sub txtOpnDat_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    Dim dtmDate As Date

    On Error GoTo Hell
    dtmDate = DateTime.Date
        
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
            If DateValue(txtOpnDat.Value) > dtmDate Then
                MsgBox prompt:="De opname datum kan niet later zijn " & _
                                "dan de huidige datum", Buttons:=vbCritical, _
                       Title:="Informedica"
                'Cancel=true
                txtOpnDat.SetFocus
                txtOpnDat.Value = vbNullString
            End If
        End If
    End If
    
    Exit Sub

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
    
    Exit Sub
    
Hell:
    Resume Next
End Sub

Private Sub UserForm_Activate()
    UserForm_Initialize
End Sub

Private Sub UserForm_Initialize()

    txtPatNum.Text = ModRange.GetRangeValue("PatNummer", "")
    
    txtOpnDat.Value = Now()
    txtANaam.Value = ModRange.GetRangeValue("_AchterNaam", "")
    txtVNaam.Value = ModRange.GetRangeValue("_VoorNaam", "")
    txtGebDat.Value = ModString.StringToDate(ModRange.GetRangeValue("GebDatum", ""))
    txtGew.Value = ModRange.GetRangeValue("Gewicht", 0) / 10
    txtLengte.Value = ModRange.GetRangeValue("Lengte", 0)
    txtWeken.Value = ModRange.GetRangeValue("_Weken", 0)
    txtDagen.Value = ModRange.GetRangeValue("_Dagen", 0)

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

