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
    
    If IsDate(txtOpnDat.Value) Then m_Pat.OpnameDatum = DateValue(txtOpnDat.Value)
    m_Pat.PatientID = txtPatNum.Value
    m_Pat.AchterNaam = txtANaam.Value
    m_Pat.VoorNaam = txtVNaam.Text
    If IsDate(txtGebDat.Value) Then m_Pat.GeboorteDatum = DateValue(txtGebDat.Value)
    m_Pat.Gewicht = Val(txtGew.Value)
    m_Pat.Lengte = Val(txtLengte.Value)
    m_Pat.GeboorteGewicht = Val(txtGebGew.Value)
    m_Pat.Weeks = Val(txtWeken.Value)
    m_Pat.Days = Val(txtDagen.Value)

    Me.Hide

End Sub

Private Sub txtDagen_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    blnCancel = Not ModPatient.ValidDagen(Val(txtDagen.Value))

End Sub

Private Sub txtDagen_KeyPress(ByVal intKey As MSForms.ReturnInteger)
    
    intKey = ModUtils.CorrectNumberAscii(intKey)

End Sub

Private Sub txtGebDat_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    If IsDate(txtGebDat.Value) Then
        If Not ModPatient.ValidBirthDate(CDate(txtGebDat.Value)) Then blnCancel = True
    Else
        blnCancel = True
    End If
    
End Sub

Public Sub SetPatient(ByRef objPat As ClassPatientDetails)

    Set m_Pat = objPat
    
    txtPatNum.Text = m_Pat.PatientID
    txtOpnDat.Value = IIf(ModDate.IsEmptyDate(m_Pat.OpnameDatum), Date, m_Pat.OpnameDatum)
    txtANaam.Value = m_Pat.AchterNaam
    txtVNaam.Value = m_Pat.VoorNaam
    txtGebDat.Value = IIf(ModDate.IsEmptyDate(m_Pat.GeboorteDatum), vbNullString, m_Pat.GeboorteDatum)
    txtGew.Value = IIf(m_Pat.Gewicht = 0, vbNullString, m_Pat.Gewicht)
    txtLengte.Value = IIf(m_Pat.Lengte = 0, vbNullString, m_Pat.Lengte)
    txtGebGew.Value = IIf(m_Pat.GeboorteGewicht = 0, vbNullString, m_Pat.GeboorteGewicht)
    txtWeken.Value = IIf(m_Pat.Weeks = 0, vbNullString, m_Pat.Weeks)
    txtDagen.Value = IIf(m_Pat.Days = 0, vbNullString, m_Pat.Days)

End Sub

Private Sub txtLengte_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    blnCancel = Not ModPatient.ValidLengthCm(txtLengte.Value)

End Sub

Private Sub txtGew_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    blnCancel = Not ModPatient.ValidWeightKg(txtGew.Value)
            
End Sub

Private Sub txtGew_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.CorrectNumberAscii(intKey)
    
End Sub

Private Sub txtLengte_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.CorrectNumberAscii(intKey)

End Sub

Private Sub txtOpnDat_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    If IsDate(txtOpnDat.Value) Then
        If Not ModPatient.ValidAdmissionDate(CDate(txtOpnDat.Value)) Then blnCancel = True
    Else
        blnCancel = True
    End If

End Sub

Private Sub txtWeken_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    blnCancel = Not ModPatient.ValidWeken(Val(txtWeken.Value))

End Sub

Private Sub txtWeken_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.CorrectNumberAscii(intKey)

End Sub
