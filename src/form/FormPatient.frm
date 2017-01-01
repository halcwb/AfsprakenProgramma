VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormPatient 
   Caption         =   "Nieuwe patient"
   ClientHeight    =   3409
   ClientLeft      =   42
   ClientTop       =   329
   ClientWidth     =   7917
   OleObjectBlob   =   "FormPatient.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Pat As ClassPatientDetails

Private Sub cmdCancel_Click()

    Me.Hide

End Sub

Private Sub cmdOk_Click()
    
    Dim dtmBd As Date
    Dim dtmAdm As Date
    
    
    If IsDate(txtOpnDat.Value) And IsDate(txtGebDat.Value) Then
        dtmBd = DateValue(txtGebDat.Value)
        dtmAdm = DateValue(txtOpnDat.Value)
        m_Pat.SetAdmissionAndBirthDate dtmAdm, dtmBd
    End If
    
    m_Pat.PatientID = txtPatNum.Value
    m_Pat.AchterNaam = txtANaam.Value
    m_Pat.VoorNaam = txtVNaam.Text
    m_Pat.Gewicht = Val(txtGew.Value)
    m_Pat.Lengte = Val(txtLengte.Value)
    m_Pat.GeboorteGewicht = Val(txtGebGew.Value)
    m_Pat.Weeks = Val(txtWeken.Value)
    m_Pat.Days = Val(txtDagen.Value)

    Me.Hide

End Sub

Private Sub txtDagen_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    If txtDagen.Value = "" Then Exit Sub

    If Not ModPatient.ValidDagen(Val(txtDagen.Value)) Then
        ModMessage.ShowMsgBoxExclam "Geen geldige waarde voor dagen: " & txtDagen.Value
        txtDagen.Value = ""
        blnCancel = True
    End If
    

End Sub

Private Sub txtDagen_KeyPress(ByVal intKey As MSForms.ReturnInteger)
    
    intKey = ModUtils.CorrectNumberAscii(intKey)

End Sub

Private Sub txtGebDat_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    Dim dtmAdm As Date
    
    If txtGebDat.Value = "" Then Exit Sub
    
    If Not IsDate(txtOpnDat.Value) Then
        dtmAdm = Date
        txtOpnDat.Value = ModString.DateToString(dtmAdm)
    Else
        dtmAdm = ModString.StringToDate((txtOpnDat.Value))
    End If

    If IsDate(txtGebDat.Value) Then
        If Not ModPatient.ValidBirthDate(ModString.StringToDate((txtGebDat.Value)), dtmAdm) Then
            ModMessage.ShowMsgBoxExclam "Geboortedatum kan niet na de opname datum liggen"
            txtGebDat.Value = ""
            blnCancel = True
        Else
            txtGebDat.Value = ModString.DateToString(txtGebDat.Value)
            blnCancel = False
        End If
    Else
        ModMessage.ShowMsgBoxExclam "Geen geldige datum: " & txtGebDat.Value
        blnCancel = True
    End If
    
End Sub

Public Sub SetPatient(ByRef objPat As ClassPatientDetails)

    Dim strOpnDat As String
    Dim strGebDat As String

    Set m_Pat = objPat
    
    strOpnDat = IIf(ModDate.IsEmptyDate(m_Pat.OpnameDatum), ModString.DateToString(Date), ModString.DateToString(m_Pat.OpnameDatum))
    strGebDat = IIf(ModDate.IsEmptyDate(m_Pat.GeboorteDatum), vbNullString, ModString.DateToString(m_Pat.GeboorteDatum))
    
    txtPatNum.Text = m_Pat.PatientID
    txtOpnDat.Value = strOpnDat
    txtANaam.Value = m_Pat.AchterNaam
    txtVNaam.Value = m_Pat.VoorNaam
    txtGebDat.Value = strGebDat
    txtGew.Value = IIf(m_Pat.Gewicht = 0, vbNullString, m_Pat.Gewicht)
    txtLengte.Value = IIf(m_Pat.Lengte = 0, vbNullString, m_Pat.Lengte)
    txtGebGew.Value = IIf(m_Pat.GeboorteGewicht = 0, vbNullString, m_Pat.GeboorteGewicht)
    txtWeken.Value = IIf(m_Pat.Weeks = 0, vbNullString, m_Pat.Weeks)
    txtDagen.Value = IIf(m_Pat.Days = 0, vbNullString, m_Pat.Days)

End Sub

Private Sub txtGebGew_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    If txtGebGew.Value = "" Then Exit Sub
    
    If Not ModPatient.ValidBirthWeight(Val(txtGebGew.Value)) Then
        ModMessage.ShowMsgBoxExclam "Geen geldig geboortegewicht: " & txtGebGew.Value
        txtGebGew.Value = ""
        blnCancel = True
    End If

End Sub

Private Sub txtGebGew_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.CorrectNumberAscii(intKey)
    
End Sub

Private Sub txtLengte_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    If txtLengte.Value = "" Then Exit Sub

    If Not ModPatient.ValidLengthCm(txtLengte.Value) Then
        ModMessage.ShowMsgBoxExclam "Geen geldige waarde voor lengte: " & txtLengte.Value
        txtLengte.Value = ""
        blnCancel = True
    End If

End Sub

Private Sub txtGew_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    If txtGew.Value = "" Then Exit Sub

    If Not ModPatient.ValidWeightKg(txtGew.Value) Then
        ModMessage.ShowMsgBoxExclam "Geen geldige waarde voor gewicht: " & txtGew.Value
        txtGew.Value = ""
        blnCancel = True
    End If
            
End Sub

Private Sub txtGew_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.CorrectNumberAscii(intKey)
    
End Sub

Private Sub txtLengte_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.CorrectNumberAscii(intKey)

End Sub

Private Sub txtOpnDat_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    Dim dtmAdm As Date
    Dim dtmGeb As Date

    If txtOpnDat.Value = "" Then Exit Sub

    If IsDate(txtOpnDat.Value) Then
        dtmAdm = ModString.StringToDate(txtOpnDat.Value)
        
        If Not ModPatient.ValidAdmissionDate(dtmAdm) Then
            ModMessage.ShowMsgBoxExclam "Geen geldige opname datum: " & txtOpnDat.Value
        
            txtOpnDat.Value = ""
            blnCancel = True
        Else
            If IsDate(txtGebDat.Value) Then
                dtmGeb = ModString.StringToDate(txtGebDat.Value)
                If Not ModPatient.ValidBirthDate(dtmGeb, dtmAdm) Then
                    ModMessage.ShowMsgBoxExclam "Opname datum kan niet voor de geboortedatum liggen"
                    txtOpnDat.Value = ""
                    blnCancel = True
                Else
                    txtOpnDat.Value = ModString.DateToString(dtmAdm)
                    blnCancel = False
                End If
            Else
                txtOpnDat.Value = ModString.DateToString(dtmAdm)
                blnCancel = False
            End If
        End If
    Else
        ModMessage.ShowMsgBoxExclam "Geen geldige datum: " & txtOpnDat.Value
        txtOpnDat.Value = ""
        blnCancel = True
    End If

End Sub

Private Sub txtWeken_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    If txtWeken.Value = "" Then Exit Sub

    If Not ModPatient.ValidWeken(Val(txtWeken.Value)) Then
        ModMessage.ShowMsgBoxExclam "Geen geldige waarde voor weken: " & txtWeken.Value
        txtWeken.Value = ""
        blnCancel = True
    End If

End Sub

Private Sub txtWeken_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.CorrectNumberAscii(intKey)

End Sub
