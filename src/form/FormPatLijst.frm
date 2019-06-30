VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormPatLijst 
   Caption         =   "Kies een patient ..."
   ClientHeight    =   11145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9345
   OleObjectBlob   =   "FormPatLijst.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormPatLijst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Pats As Collection
Private m_OriginalPats As Collection
Private m_StandardPats As Collection
Private m_Versions As Collection
Private m_OnlyAdmitted As Boolean
Private m_LatestVersion As Boolean
Private m_UseDatabase As Boolean
Private m_Cancel As Boolean
Private m_Standard As Boolean

Public Function GetCancel() As Boolean

    GetCancel = m_Cancel

End Function

Private Sub SetSelectedHospNumAndVersion()

    If Not m_LatestVersion Then
        If cboVersions.Value = vbNullString Then
            ModMessage.ShowMsgBoxExclam "Selecteer eerst een afspraken versie"
            Exit Sub
        Else
            ModBed.SetDatabaseVersie Database_GetVersionIDFromString(cboVersions.Value)
        End If
    End If
    
    If m_UseDatabase Then
        SetSelectedHospNum
    Else
        SetSelectedBed
    End If


End Sub

Private Sub cmdCancel_Click()

    m_Cancel = True
    Me.Hide

End Sub

Private Sub cmdOK_Click()

    SetSelectedHospNumAndVersion
    m_Cancel = False
    Me.Hide

End Sub

Private Sub lstPatienten_Click()

    If Not m_LatestVersion Then
        LoadVersions
    End If

End Sub

Private Sub lstPatienten_DblClick(ByVal blnCancel As MSForms.ReturnBoolean)
        
    SetSelectedVersion
    If m_UseDatabase Then SetSelectedHospNum
    Me.Hide

End Sub

Public Sub LoadPatients(ByVal colPats As Collection)

    Dim objPat As ClassPatientInfo
    Set m_Pats = colPats
    
    For Each objPat In colPats
        Me.lstPatienten.AddItem objPat.ToString
    Next objPat

End Sub

Function SortArrayAtoZ(arrArray As Variant)

    Dim lngI As Long
    Dim lngJ As Long
    Dim lngC As Long
    Dim varTemp
    
    'Sort the Array A-Z
    'ModProgress.StartProgress "Sorteren van patienten"
    
    For lngI = LBound(arrArray) To UBound(arrArray) - 1
        For lngJ = lngI + 1 To UBound(arrArray)
            If UCase(arrArray(lngI)) > UCase(arrArray(lngJ)) Then
                varTemp = arrArray(lngJ)
                arrArray(lngJ) = arrArray(lngI)
                arrArray(lngI) = varTemp
            End If
        Next lngJ
        'ModProgress.SetJobPercentage "Sorteren", lngC, lngI
    Next lngI
    
    SortArrayAtoZ = arrArray

End Function

Private Function PatToSortString(objPat As ClassPatientDetails) As String

    Dim strSort As String
    
    strSort = objPat.AchterNaam & objPat.VoorNaam & objPat.HospitalNumber
    strSort = IIf(m_OnlyAdmitted, objPat.Bed & strSort, strSort)
    
    PatToSortString = strSort

End Function

Private Sub UseDatabase()

    m_UseDatabase = True
    frmPatSel.Visible = True
    frmVersion.Visible = True

End Sub

Public Sub LoadDbPatients(ByVal colPats As Collection, ByVal blnStandard As Boolean)

    Dim objPat As ClassPatientDetails
    Dim arrSort() As Variant
    Dim varSort As Variant
    Dim strPat As String
    
    UseDatabase
   
    Set m_Pats = New Collection
    
    If blnStandard Then
        Set m_StandardPats = colPats
    Else
        Set m_OriginalPats = colPats
    End If
    
    Me.lstPatienten.Clear
    
    For Each objPat In colPats
        If m_OnlyAdmitted And Not blnStandard Then
            If Not objPat.Bed = vbNullString And objPat.Afdeling = ModMetaVision.MetaVision_GetDepartment() Then
                ModArray.AddItemToVariantArray arrSort, PatToSortString(objPat)
            End If
        Else
            ModArray.AddItemToVariantArray arrSort, PatToSortString(objPat)
        End If
    Next objPat
    
    If colPats.Count > 0 Then
        arrSort = SortArrayAtoZ(arrSort)
        
        For Each varSort In arrSort
            For Each objPat In colPats
                If PatToSortString(objPat) = varSort Then
                    m_Pats.Add objPat
                    strPat = objPat.ToString
                    strPat = IIf(m_OnlyAdmitted, objPat.Bed & " - " & strPat, strPat)
                    
                    Me.lstPatienten.AddItem strPat
                End If
            Next objPat
        Next varSort
    End If
    
End Sub

Private Function GetSelectedHospNum() As String

    
    Dim objPat As ClassPatientDetails
    Dim strId As String

    If Me.lstPatienten.ListIndex > -1 Then
        Set objPat = m_Pats(Me.lstPatienten.ListIndex + 1)
        strId = objPat.HospitalNumber
    End If

    GetSelectedHospNum = strId

End Function

Private Sub SetSelectedHospNum()
    
    Dim objPat As ClassPatientDetails
    Dim strId As String

    strId = GetSelectedHospNum()
    
    Patient_SetHospitalNumber strId

End Sub

Private Sub SetSelectedVersion()

    ModBed.SetDatabaseVersie IIf(cboVersions.Value = vbNullString, 0, cboVersions.Value)

End Sub

Public Sub SetSelectedBed()
    
    Dim objPat As ClassPatientInfo
    Dim strBed As String

    If Me.lstPatienten.ListIndex > -1 Then
        Set objPat = m_Pats(Me.lstPatienten.ListIndex + 1)
        strBed = objPat.Bed
    End If
    
    ModBed.SetBed strBed

End Sub

Private Sub CenterForm()

    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

End Sub

Private Sub ToggleAdmitted(ByVal blnAdmitted As Boolean)

    m_OnlyAdmitted = blnAdmitted
    m_Standard = False
    If Not m_OriginalPats Is Nothing Then LoadDbPatients m_OriginalPats, False

End Sub

Private Sub optAdmitted_Click()

    ToggleAdmitted optAdmitted.Value

End Sub

Private Sub optAllPatients_Click()

    ToggleAdmitted Not optAllPatients.Value

End Sub

Private Sub optStandard_Click()

    m_Standard = True
    m_OnlyAdmitted = False
    If Not m_StandardPats Is Nothing Then LoadDbPatients m_StandardPats, True

End Sub

Private Sub optLatest_Click()
    
    SetLatestVersionTrue

End Sub

Private Sub optSpecific_Click()

    SetSpecificVersionTrue

End Sub

Private Sub UserForm_Activate()

    CenterForm

End Sub

Private Sub SetCboVersionsVisible(ByVal blnVisible As Boolean)

    lblCboVersions.Visible = blnVisible
    cboVersions.Visible = blnVisible

End Sub


Private Sub UserForm_Initialize()

    SetOnlyAdmittedTrue
    SetLatestVersionTrue
    
    frmPatSel.Visible = False
    frmVersion.Visible = False

End Sub

Public Sub SetOnlyAdmittedTrue()

    m_OnlyAdmitted = True
    optAdmitted.Value = True

End Sub

Public Sub SetLatestVersionTrue()

    m_LatestVersion = True
    optLatest.Value = True
    SetCboVersionsVisible False
    cboVersions.Clear
    
End Sub

Public Sub SetSpecificVersionTrue()

    m_LatestVersion = False
    optSpecific.Value = True
    SetCboVersionsVisible True
    LoadVersions

End Sub

Private Sub LoadVersions()

    Dim strHospNum
    Dim objVersion As ClassVersion
    
    cboVersions.Clear
    strHospNum = GetSelectedHospNum()
    If Not strHospNum = vbNullString Then
        Set m_Versions = ModDatabase.Database_GetDataVersions(strHospNum)
        
        For Each objVersion In m_Versions
            cboVersions.AddItem objVersion.ToString()
        Next
    End If

End Sub
