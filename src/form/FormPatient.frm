VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormPatient 
   Caption         =   "Nieuwe patient"
   ClientHeight    =   3633
   ClientLeft      =   42
   ClientTop       =   329
   ClientWidth     =   8694
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

Private Sub ClearBirthDate()

    txtBirthYear.Value = vbNullString
    txtBirthMonth.Value = vbNullString
    txtBirthDay.Value = vbNullString

End Sub

Private Function BirthDateComplete() As Boolean

    BirthDateComplete = txtBirthDay.Value <> vbNullString And txtBirthMonth.Value <> vbNullString And txtBirthYear.Value <> vbNullString

End Function

Private Function GetBirthDate() As Date

    If BirthDateComplete() Then
        GetBirthDate = DateSerial(Int(txtBirthYear.Value), Int(txtBirthMonth.Value), Int(txtBirthDay.Value))
    Else
        GetBirthDate = ModDate.EmptyDate
    End If

End Function

Private Sub SetBirthDate(ByVal dtmDate As Date)

    txtBirthYear.Value = Year(dtmDate)
    txtBirthMonth.Value = Month(dtmDate)
    txtBirthDay.Value = Day(dtmDate)

End Sub

Private Sub ClearAdmDate()

    txtAdmYear.Value = vbNullString
    txtAdmMonth.Value = vbNullString
    txtAdmDay.Value = vbNullString

End Sub

Private Function AdmDateComplete() As Boolean

    AdmDateComplete = txtAdmDay.Value <> vbNullString And txtAdmMonth.Value <> vbNullString And txtAdmYear.Value <> vbNullString

End Function

Private Function GetAdmDate() As Date

    If AdmDateComplete() Then
        GetAdmDate = DateSerial(Int(txtAdmYear.Value), Int(txtAdmMonth.Value), Int(txtAdmDay.Value))
    Else
        GetAdmDate = ModDate.EmptyDate
    End If

End Function

Private Sub SetAdmDate(ByVal dtmDate As Date)

    txtAdmYear.Value = Year(dtmDate)
    txtAdmMonth.Value = Month(dtmDate)
    txtAdmDay.Value = Day(dtmDate)

End Sub

Private Sub cmdOk_Click()
    
    Dim dtmBd As Date
    Dim dtmAdm As Date
    
    dtmAdm = GetAdmDate()
    dtmBd = GetBirthDate()
    If Not ModDate.IsEmptyDate(dtmAdm) And Not ModDate.IsEmptyDate(dtmBd) Then
        m_Pat.SetAdmissionAndBirthDate dtmAdm, dtmBd
    End If
    
    m_Pat.PatientID = txtPatNum.Value
    m_Pat.AchterNaam = txtLastName.Value
    m_Pat.VoorNaam = txtFirstName.Text
    m_Pat.Gewicht = Val(txtWeight.Value)
    m_Pat.Lengte = Val(txtLength.Value)
    m_Pat.GeboorteGewicht = Val(txtBirthWeight.Value)
    m_Pat.Weeks = Val(txtGestWeek.Value)
    m_Pat.Days = Val(txtGestDay.Value)

    Me.Hide

End Sub

Private Sub txtBirthYear_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    If BirthDateComplete() Then blnCancel = Not ValidateBirthDate()

End Sub

Private Sub txtBirthMonth_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    If BirthDateComplete() Then blnCancel = Not ValidateBirthDate()

End Sub

Private Sub txtBirthDay_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    If BirthDateComplete() Then blnCancel = Not ValidateBirthDate()

End Sub

Private Sub txtAdmYear_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    If AdmDateComplete() Then blnCancel = Not ValidateAdmDate()

End Sub

Private Sub txtAdmMonth_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    If AdmDateComplete() Then blnCancel = Not ValidateAdmDate()

End Sub

Private Sub txtAdmDay_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    If AdmDateComplete() Then blnCancel = Not ValidateAdmDate()

End Sub

Private Sub txtGestDay_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    If txtGestDay.Value = vbNullString Then Exit Sub

    If Not ModPatient.ValidDagen(Val(txtGestDay.Value)) Then
        ModMessage.ShowMsgBoxExclam "Geen geldige waarde voor dagen: " & txtGestDay.Value
        txtGestDay.Value = vbNullString
        blnCancel = True
    End If

End Sub

Private Sub txtGestDay_KeyPress(ByVal intKey As MSForms.ReturnInteger)
    
    intKey = ModUtils.CorrectNumberAscii(intKey)

End Sub

Private Function ValidateBirthDate() As Boolean

    Dim dtmAdm As Date
    Dim dtmBd As Date
    
    dtmAdm = GetAdmDate()
    dtmBd = GetBirthDate()
    
    If ModDate.IsEmptyDate(dtmBd) Then Exit Function
    
    If ModDate.IsEmptyDate(dtmAdm) Then
        dtmAdm = DateTime.Date
        SetAdmDate (dtmAdm)
    End If

    If Not ModPatient.ValidBirthDate(dtmBd, dtmAdm) Then
        ModMessage.ShowMsgBoxExclam "Geboortedatum kan niet na de opname datum liggen"
        
        ClearBirthDate
        ValidateBirthDate = False
    Else
        ValidateBirthDate = True
    End If
    
End Function

Public Sub SetPatient(ByRef objPat As ClassPatientDetails)

    Set m_Pat = objPat
    
    If Not ModDate.IsEmptyDate(m_Pat.OpnameDatum) Then
        SetAdmDate m_Pat.OpnameDatum
    Else
        SetAdmDate DateTime.Date
    End If
    If Not ModDate.IsEmptyDate(m_Pat.GeboorteDatum) Then SetBirthDate m_Pat.GeboorteDatum
    
    txtPatNum.Text = m_Pat.PatientID
    txtLastName.Value = m_Pat.AchterNaam
    txtFirstName.Value = m_Pat.VoorNaam
    txtWeight.Value = IIf(m_Pat.Gewicht = 0, vbNullString, m_Pat.Gewicht)
    txtLength.Value = IIf(m_Pat.Lengte = 0, vbNullString, m_Pat.Lengte)
    txtBirthWeight.Value = IIf(m_Pat.GeboorteGewicht = 0, vbNullString, m_Pat.GeboorteGewicht)
    txtGestWeek.Value = IIf(m_Pat.Weeks = 0, vbNullString, m_Pat.Weeks)
    txtGestDay.Value = IIf(m_Pat.Days = 0, vbNullString, m_Pat.Days)

End Sub

Private Sub txtBirthWeight_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    If txtBirthWeight.Value = vbNullString Then Exit Sub
    
    If Not ModPatient.ValidBirthWeight(Val(txtBirthWeight.Value)) Then
        ModMessage.ShowMsgBoxExclam "Geen geldig geboortegewicht: " & txtBirthWeight.Value
        txtBirthWeight.Value = vbNullString
        blnCancel = True
    End If

End Sub

Private Sub txtBirthWeight_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.OnlyNumericAscii(intKey)
    
End Sub

Private Sub txtLength_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    If txtLength.Value = vbNullString Then Exit Sub

    If Not ModPatient.ValidLengthCm(txtLength.Value) Then
        ModMessage.ShowMsgBoxExclam "Geen geldige waarde voor lengte: " & txtLength.Value
        txtLength.Value = vbNullString
        blnCancel = True
    End If

End Sub

Private Sub txtWeight_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    If txtWeight.Value = vbNullString Then Exit Sub

    If Not ModPatient.ValidWeightKg(txtWeight.Value) Then
        ModMessage.ShowMsgBoxExclam "Geen geldige waarde voor gewicht: " & txtWeight.Value
        txtWeight.Value = vbNullString
        blnCancel = True
    End If
            
End Sub

Private Sub txtWeight_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.CorrectNumberAscii(intKey)
    
End Sub

Private Sub txtLength_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.OnlyNumericAscii(intKey)

End Sub

Private Function ValidateAdmDate() As Boolean

    Dim strText As String

    Dim dtmAdm As Date
    Dim dtmBd As Date
      
    dtmAdm = GetAdmDate()
        
    If Not ModPatient.ValidAdmissionDate(dtmAdm) Then
        ModMessage.ShowMsgBoxExclam "Geen geldige opname datum: " & ModString.DateToString(dtmAdm)
        
        ClearAdmDate
        ValidateAdmDate = False
    Else
        If BirthDateComplete() Then
            dtmBd = GetBirthDate()
            If Not ModPatient.ValidBirthDate(dtmBd, dtmAdm) Then
                strText = "Opname datum " & ModString.DateToString(dtmAdm)
                strText = strText & " kan niet voor geboortedatum " & ModString.DateToString(dtmBd) & " liggen"
                ModMessage.ShowMsgBoxExclam "Opname datum kan niet voor de geboortedatum liggen"
                ClearAdmDate
                
                ValidateAdmDate = False
            Else
                ValidateAdmDate = True
            End If
        Else
            ValidateAdmDate = True
        End If
    End If

End Function

Private Sub txtGestWeek_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    If txtGestWeek.Value = vbNullString Then Exit Sub

    If Not ModPatient.ValidWeken(Val(txtGestWeek.Value)) Then
        ModMessage.ShowMsgBoxExclam "Geen geldige waarde voor weken: " & txtGestWeek.Value
        txtGestWeek.Value = vbNullString
        blnCancel = True
    End If

End Sub

Private Sub txtGestWeek_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.OnlyNumericAscii(intKey)

End Sub

Private Sub txtBirthDay_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.OnlyNumericAscii(intKey)

End Sub

Private Sub txtBirthMonth_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.OnlyNumericAscii(intKey)

End Sub

Private Sub txtBirthYear_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.OnlyNumericAscii(intKey)

End Sub

Private Sub txtAdmDay_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.OnlyNumericAscii(intKey)

End Sub

Private Sub txtAdmMonth_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.OnlyNumericAscii(intKey)

End Sub

Private Sub txtAdmYear_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.OnlyNumericAscii(intKey)

End Sub
