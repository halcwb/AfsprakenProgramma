VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormPatient 
   Caption         =   "Nieuwe patient"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9525
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
            strValid = IIf(txtGestWeek.Value = vbNullString, "Voer zwangerschaps duur in", strValid)
            strValid = IIf(txtBirthWeight.Value = vbNullString, "Voer geboortegewicht in", strValid)
        End If
    End If
    
    strValid = IIf(cboGeslacht.Text = vbNullString, "Voer geslacht in", strValid)
    
    strValid = IIf(txtLength.Value = vbNullString, "Voer lengte in", strValid)
    strValid = IIf(txtWeight.Value = vbNullString, "Voer gewicht in", strValid)
    strValid = IIf(Not BirthDateComplete, "Voer geboorte datum in", strValid)
    strValid = IIf(Not AdmDateComplete, "Voer opname datum in", strValid)
    
    strValid = IIf(txtFirstName.Value = vbNullString, "Vul voor naam  in", strValid)
    strValid = IIf(txtLastName.Value = vbNullString, "Vul achter naam  in", strValid)
    strValid = IIf(txtPatNum.Value = vbNullString, "Vul patient nummer in", strValid)
    
    strValid = IIf(strText = vbNullString, strValid, strText)
    lblValid.Caption = strValid
    cmdOK.Enabled = strValid = vbNullString

End Sub

Private Sub btnAdmNow_Click()
    
    Dim dtmNow As Date
    
    dtmNow = DateTime.Date
    txtAdmDay.Value = DateTime.Day(dtmNow)
    txtAdmMonth.Value = DateTime.Month(dtmNow)
    txtAdmYear.Value = DateTime.Year(dtmNow)

End Sub

Private Sub btnBdNow_Click()

    Dim dtmNow As Date
    
    dtmNow = DateTime.Date
    txtBirthDay.Value = DateTime.Day(dtmNow)
    txtBirthMonth.Value = DateTime.Month(dtmNow)
    txtBirthYear.Value = DateTime.Year(dtmNow)

End Sub

Private Sub btnRefresh_Click()

    Dim strId As String
    
    strId = IIf(txtPatNum.Text = vbNullString, MetaVision_GetCurrentPatientID(), vbNullString)
    If Not (strId = vbNullString And txtPatNum.Text = vbNullString) Then
        MetaVision_GetPatientDetails m_Pat, strId, txtPatNum.Text
        Me.txtAdmDay = DateTime.Day(m_Pat.OpnameDatum)
        Me.txtAdmMonth = DateTime.Month(m_Pat.OpnameDatum)
        Me.txtAdmYear = DateTime.Year(m_Pat.OpnameDatum)
        Me.txtPatNum = m_Pat.PatientId
        Me.txtLastName = m_Pat.AchterNaam
        Me.txtFirstName = m_Pat.VoorNaam
        Me.txtBirthDay = DateTime.Day(m_Pat.GeboorteDatum)
        Me.txtBirthMonth = DateTime.Month(m_Pat.GeboorteDatum)
        Me.txtBirthYear = DateTime.Year(m_Pat.GeboorteDatum)
        Me.txtWeight = m_Pat.Gewicht
        Me.txtLength = m_Pat.Lengte
        Me.cboGeslacht.Text = m_Pat.Geslacht
        Me.txtBirthWeight = m_Pat.GeboorteGewicht
        Me.txtGestDay = m_Pat.Days
        Me.txtGestWeek = m_Pat.Weeks
        
        Validate vbNullString
    End If
    
    ModMetaVision.MetaVision_SyncLab

End Sub

Private Sub cboGeslacht_Change()

    If cboGeslacht.ListIndex = -1 Then
        cboGeslacht.SetFocus
        ModMessage.ShowMsgBoxInfo "Geef een geldig geslacht op"
    End If

    Validate vbNullString

End Sub

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

Private Sub cmdClear_Click()

    txtAdmDay.Value = vbNullString
    txtAdmMonth.Value = vbNullString
    txtAdmYear.Value = vbNullString
    
    txtPatNum.Value = vbNullString
    txtLastName.Value = vbNullString
    txtFirstName.Value = vbNullString
    
    txtBirthDay.Value = vbNullString
    txtBirthMonth.Value = vbNullString
    txtBirthYear.Value = vbNullString
    
    txtWeight.Value = vbNullString
    txtLength.Value = vbNullString
    
    cboGeslacht.Value = vbNullString
    
    txtBirthWeight.Value = vbNullString
    txtGestWeek.Value = vbNullString
    txtGestDay.Value = vbNullString

End Sub

Private Sub cmdOK_Click()
    
    Dim dtmBD As Date
    Dim dtmAdm As Date
    Dim dblActWght As Double
    Dim dblBthWght As Double
    
    dtmAdm = GetAdmDate()
    dtmBD = GetBirthDate()
    If Not ModDate.IsEmptyDate(dtmAdm) And Not ModDate.IsEmptyDate(dtmBD) Then
        m_Pat.SetAdmissionAndBirthDate dtmAdm, dtmBD
    End If
    
    dblActWght = StringToDouble(txtWeight.Value)
    dblBthWght = StringToDouble(txtBirthWeight.Value)
    If dblActWght < dblBthWght / 1000 Then dblActWght = dblBthWght / 1000
    
    m_Pat.PatientId = CStr(txtPatNum.Text)
    m_Pat.AchterNaam = txtLastName.Text
    m_Pat.VoorNaam = txtFirstName.Text
    m_Pat.Gewicht = dblActWght
    m_Pat.Lengte = ModString.StringToDouble(txtLength.Value)
    m_Pat.Geslacht = cboGeslacht.Text
    m_Pat.GeboorteGewicht = dblBthWght
    m_Pat.Weeks = StringToDouble(txtGestWeek.Value)
    m_Pat.Days = StringToDouble(txtGestDay.Value)
    
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

    If txtGestDay.Value = vbNullString Then Exit Sub

    If Not ModPatient.ValidDagen(StringToDouble(txtGestDay.Value)) Then
        txtGestDay.Value = vbNullString
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
    Dim dtmBD As Date
    
    dtmAdm = GetAdmDate()
    dtmBD = GetBirthDate()
    
    If ModDate.IsEmptyDate(dtmBD) Then Exit Function
    
    If ModDate.IsEmptyDate(dtmAdm) Then
        dtmAdm = DateTime.Date
        SetAdmDate (dtmAdm)
    End If

    If Not ModPatient.ValidBirthDate(dtmBD, dtmAdm) Then
        strValid = "Geboortedatum na opname datum"
        Validate strValid
        
        ClearBirthDate
        ValidateBirthDate = False
    Else
        SetBirthDate GetBirthDate
        Validate vbNullString
        ValidateBirthDate = True
    End If
    
End Function

Public Sub SetPatient(objPat As ClassPatientDetails)

    Set m_Pat = objPat
    
    If Not ModDate.IsEmptyDate(m_Pat.OpnameDatum) Then
        SetAdmDate m_Pat.OpnameDatum
    Else
        SetAdmDate DateTime.Date
    End If
    If Not ModDate.IsEmptyDate(m_Pat.GeboorteDatum) Then SetBirthDate m_Pat.GeboorteDatum
    
    txtPatNum.Text = m_Pat.PatientId
    txtLastName.Value = m_Pat.AchterNaam
    txtFirstName.Value = m_Pat.VoorNaam
    txtWeight.Value = IIf(m_Pat.Gewicht = 0, vbNullString, m_Pat.Gewicht)
    txtLength.Value = IIf(m_Pat.Lengte = 0, vbNullString, m_Pat.Lengte)
    cboGeslacht.Text = objPat.Geslacht
    txtBirthWeight.Value = IIf(m_Pat.GeboorteGewicht = 0, vbNullString, m_Pat.GeboorteGewicht)
    txtGestWeek.Value = IIf(m_Pat.Weeks = 0, vbNullString, m_Pat.Weeks)
    txtGestDay.Value = IIf(m_Pat.Days = 0, vbNullString, m_Pat.Days)

End Sub

Private Sub txtBirthWeight_BeforeUpdate(ByVal blnCancel As MSForms.ReturnBoolean)

    Dim strValid As String
    
    If txtBirthWeight.Value = vbNullString Then Exit Sub
    
    If Not ModPatient.ValidBirthWeight(StringToDouble(txtBirthWeight.Value)) Then
        strValid = "Geen geldig geboortegewicht"
        Validate strValid
        
        txtBirthWeight.Value = vbNullString
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
    
    If txtLength.Value = vbNullString Then Exit Sub

    If Not ModPatient.ValidLengthCm(txtLength.Value) Then
        strValid = "Geen geldige lengte"
        Validate strValid
        
        txtLength.Value = vbNullString
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

    If txtWeight.Value = vbNullString Then Exit Sub

    If Not ModPatient.ValidWeightKg(txtWeight.Value) Then
        strValid = "Geen geldige gewicht"
        Validate strValid
        
        txtWeight.Value = vbNullString
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
    Dim dtmBD As Date
      
    dtmAdm = GetAdmDate()
        
    If Not ModPatient.ValidAdmissionDate(dtmAdm) Then
        strValid = "Geen geldige opname datum"
        Validate strValid
        
        ClearAdmDate
        ValidateAdmDate = False
    Else
        If BirthDateComplete() Then
            dtmBD = GetBirthDate()
            If Not ModPatient.ValidBirthDate(dtmBD, dtmAdm) Then
                strValid = "Opname datum voor geboorte datum"
                Validate strValid
                
                ClearAdmDate
                ValidateAdmDate = False
            Else
                SetAdmDate GetAdmDate
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

    If txtGestWeek.Value = vbNullString Then Exit Sub

    If Not ModPatient.ValidWeken(StringToDouble(txtGestWeek.Value)) Then
        strValid = "Geen zwangerschapsduur"
        Validate strValid
        
        txtGestWeek.Value = vbNullString
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

Private Sub CenterForm()

    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

End Sub

Private Sub UserForm_Activate()

    CenterForm
    
    Validate vbNullString
    
End Sub

Private Sub UserForm_Initialize()

    Me.txtAdmDay.TabIndex = 1
    Me.txtAdmMonth.TabIndex = 2
    Me.txtAdmYear.TabIndex = 3
    
    Me.txtPatNum.TabIndex = 4
    
    Me.txtLastName.TabIndex = 5
    Me.txtFirstName.TabIndex = 6
    
    Me.txtBirthDay.TabIndex = 7
    Me.txtBirthMonth.TabIndex = 8
    Me.txtBirthYear.TabIndex = 9
    
    Me.txtWeight.TabIndex = 10
    Me.txtLength.TabIndex = 11
    
    Me.cboGeslacht.TabIndex = 12
    
    Me.txtGestWeek.TabIndex = 13
    Me.txtGestDay.TabIndex = 14
    
    Me.txtBirthWeight.TabIndex = 15
    
    Me.cmdOK.TabIndex = 16
    Me.cmdClear.TabIndex = 17
    Me.cmdCancel.TabIndex = 18
    
    cboGeslacht.AddItem "man"
    cboGeslacht.AddItem "vrouw"
    cboGeslacht.AddItem "onbepaald"

End Sub
