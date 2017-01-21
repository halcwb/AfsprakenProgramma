VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormPatient 
   Caption         =   "Nieuwe patient"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9045
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

Private Sub Validate(ByVal strText As String)

    Dim strValid As String
    
    strValid = vbNullString
    
    If BirthDateComplete() Then
        If DateTime.DateDiff("d", GetBirthDate(), DateTime.Date) <= 28 Then
            strValid = IIf(txtGestWeek.value = vbNullString, "Voer zwangerschaps duur in", strValid)
            strValid = IIf(txtBirthWeight.value = vbNullString, "Voer geboortegewicht in", strValid)
        End If
    End If
    
    strValid = IIf(txtLength.value = vbNullString, "Voer lengte in", strValid)
    strValid = IIf(txtWeight.value = vbNullString, "Voer gewicht in", strValid)
    strValid = IIf(Not BirthDateComplete, "Voer geboorte datum in", strValid)
    strValid = IIf(Not AdmDateComplete, "Voer opname datum in", strValid)
    
    strValid = IIf(txtFirstName.value = vbNullString, "Vul voor naam  in", strValid)
    strValid = IIf(txtLastName.value = vbNullString, "Vul achter naam  in", strValid)
    strValid = IIf(txtPatNum.value = vbNullString, "Vul patient nummer in", strValid)
    
    strValid = IIf(strText = vbNullString, strValid, strText)
    lblValid.Caption = strValid
    cmdOK.Enabled = strValid = vbNullString

End Sub

Private Sub btnNow_Click()
    
    Dim dtmNow As Date
    
    dtmNow = DateTime.Date
    txtAdmDay.value = DateTime.Day(dtmNow)
    txtAdmMonth.value = DateTime.Month(dtmNow)
    txtAdmYear.value = DateTime.Year(dtmNow)

End Sub

Private Sub cmdCancel_Click()

    Me.Hide

End Sub

Private Sub ClearBirthDate()

    txtBirthYear.value = vbNullString
    txtBirthMonth.value = vbNullString
    txtBirthDay.value = vbNullString

End Sub

Private Function BirthDateComplete() As Boolean

    BirthDateComplete = txtBirthDay.value <> vbNullString And txtBirthMonth.value <> vbNullString And txtBirthYear.value <> vbNullString

End Function

Private Function GetBirthDate() As Date

    If BirthDateComplete() Then
        GetBirthDate = DateSerial(Int(txtBirthYear.value), Int(txtBirthMonth.value), Int(txtBirthDay.value))
    Else
        GetBirthDate = ModDate.EmptyDate
    End If

End Function

Private Sub SetBirthDate(ByVal dtmDate As Date)

    txtBirthYear.value = Year(dtmDate)
    txtBirthMonth.value = Month(dtmDate)
    txtBirthDay.value = Day(dtmDate)

End Sub

Private Sub ClearAdmDate()

    txtAdmYear.value = vbNullString
    txtAdmMonth.value = vbNullString
    txtAdmDay.value = vbNullString

End Sub

Private Function AdmDateComplete() As Boolean

    AdmDateComplete = txtAdmDay.value <> vbNullString And txtAdmMonth.value <> vbNullString And txtAdmYear.value <> vbNullString

End Function

Private Function GetAdmDate() As Date

    If AdmDateComplete() Then
        GetAdmDate = DateSerial(Int(txtAdmYear.value), Int(txtAdmMonth.value), Int(txtAdmDay.value))
    Else
        GetAdmDate = ModDate.EmptyDate
    End If

End Function

Private Sub SetAdmDate(ByVal dtmDate As Date)

    txtAdmYear.value = Year(dtmDate)
    txtAdmMonth.value = Month(dtmDate)
    txtAdmDay.value = Day(dtmDate)

End Sub

Private Sub cmdClear_Click()

    txtAdmDay.value = vbNullString
    txtAdmMonth.value = vbNullString
    txtAdmYear.value = vbNullString
    
    txtPatNum.value = vbNullString
    txtLastName.value = vbNullString
    txtFirstName.value = vbNullString
    
    txtBirthDay.value = vbNullString
    txtBirthMonth.value = vbNullString
    txtBirthYear.value = vbNullString
    
    txtWeight.value = vbNullString
    txtLength.value = vbNullString
    
    txtBirthWeight.value = vbNullString
    txtGestWeek.value = vbNullString
    txtGestDay.value = vbNullString

End Sub

Private Sub cmdOK_Click()
    
    Dim dtmBd As Date
    Dim dtmAdm As Date
    
    dtmAdm = GetAdmDate()
    dtmBd = GetBirthDate()
    If Not ModDate.IsEmptyDate(dtmAdm) And Not ModDate.IsEmptyDate(dtmBd) Then
        m_Pat.SetAdmissionAndBirthDate dtmAdm, dtmBd
    End If
    
    m_Pat.PatientID = txtPatNum.value
    m_Pat.AchterNaam = txtLastName.value
    m_Pat.VoorNaam = txtFirstName.Text
    m_Pat.Gewicht = val(txtWeight.value)
    m_Pat.Lengte = val(txtLength.value)
    m_Pat.GeboorteGewicht = val(txtBirthWeight.value)
    m_Pat.Weeks = val(txtGestWeek.value)
    m_Pat.Days = val(txtGestDay.value)

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

Private Sub txtFirstName_Change()

    Validate vbNullString

End Sub

Private Sub txtGestDay_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    Dim strValid As String

    If txtGestDay.value = vbNullString Then Exit Sub

    If Not ModPatient.ValidDagen(val(txtGestDay.value)) Then
        txtGestDay.value = vbNullString
        blnCancel = True
    End If
    
    Validate vbNullString

End Sub

Private Sub txtGestDay_KeyPress(ByVal intKey As MSForms.ReturnInteger)
    
    intKey = ModUtils.CorrectNumberAscii(intKey)

End Sub

Private Function ValidateBirthDate() As Boolean

    Dim strValid As String
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
        strValid = "Geboortedatum na opname datum"
        Validate strValid
        
        ClearBirthDate
        ValidateBirthDate = False
    Else
        Validate vbNullString
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
    txtLastName.value = m_Pat.AchterNaam
    txtFirstName.value = m_Pat.VoorNaam
    txtWeight.value = IIf(m_Pat.Gewicht = 0, vbNullString, m_Pat.Gewicht)
    txtLength.value = IIf(m_Pat.Lengte = 0, vbNullString, m_Pat.Lengte)
    txtBirthWeight.value = IIf(m_Pat.GeboorteGewicht = 0, vbNullString, m_Pat.GeboorteGewicht)
    txtGestWeek.value = IIf(m_Pat.Weeks = 0, vbNullString, m_Pat.Weeks)
    txtGestDay.value = IIf(m_Pat.Days = 0, vbNullString, m_Pat.Days)

End Sub

Private Sub txtBirthWeight_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    Dim strValid As String
    
    If txtBirthWeight.value = vbNullString Then Exit Sub
    
    If Not ModPatient.ValidBirthWeight(val(txtBirthWeight.value)) Then
        strValid = "Geen geldig geboortegewicht"
        Validate strValid
        
        txtBirthWeight.value = vbNullString
        blnCancel = True
    Else
        Validate vbNullString
        blnCancel = False
    End If

End Sub

Private Sub txtBirthWeight_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.OnlyNumericAscii(intKey)
    
End Sub

Private Sub txtLastName_Change()

    Validate vbNullString

End Sub

Private Sub txtLength_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    Dim strValid As String
    
    If txtLength.value = vbNullString Then Exit Sub

    If Not ModPatient.ValidLengthCm(txtLength.value) Then
        strValid = "Geen geldige lengte"
        Validate strValid
        
        txtLength.value = vbNullString
        blnCancel = True
    Else
        Validate vbNullString
        blnCancel = False
    End If

End Sub

Private Sub txtPatNum_Change()

    Validate vbNullString

End Sub

Private Sub txtWeight_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    Dim strValid As String

    If txtWeight.value = vbNullString Then Exit Sub

    If Not ModPatient.ValidWeightKg(txtWeight.value) Then
        strValid = "Geen geldige gewicht"
        Validate strValid
        
        txtWeight.value = vbNullString
        blnCancel = True
    Else
        Validate vbNullString
        blnCancel = False
    End If
    
            
End Sub

Private Sub txtWeight_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.CorrectNumberAscii(intKey)
    
End Sub

Private Sub txtLength_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.OnlyNumericAscii(intKey)

End Sub

Private Function ValidateAdmDate() As Boolean

    Dim strValid As String

    Dim dtmAdm As Date
    Dim dtmBd As Date
      
    dtmAdm = GetAdmDate()
        
    If Not ModPatient.ValidAdmissionDate(dtmAdm) Then
        strValid = "Geen geldige opname datum"
        Validate strValid
        
        ClearAdmDate
        ValidateAdmDate = False
    Else
        If BirthDateComplete() Then
            dtmBd = GetBirthDate()
            If Not ModPatient.ValidBirthDate(dtmBd, dtmAdm) Then
                strValid = "Opname datum voor geboorte datum"
                Validate strValid
                
                ClearAdmDate
                ValidateAdmDate = False
            Else
                Validate vbNullString
                ValidateAdmDate = True
            End If
        Else
            Validate vbNullString
            ValidateAdmDate = True
        End If
    End If
    

End Function

Private Sub txtGestWeek_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    Dim strValid As String

    If txtGestWeek.value = vbNullString Then Exit Sub

    If Not ModPatient.ValidWeken(val(txtGestWeek.value)) Then
        strValid = "Geen zwangerschapsduur"
        Validate strValid
        
        txtGestWeek.value = vbNullString
        blnCancel = True
    Else
        Validate vbNullString
        blnCancel = False
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

Private Sub UserForm_Activate()

    Validate vbNullString
    
End Sub
