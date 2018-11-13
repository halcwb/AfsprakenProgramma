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

Private m_onlyAdmitted As Boolean
Private m_useDatabase As Boolean

Private Sub chkAdmitted_Click()

    m_onlyAdmitted = chkAdmitted.Value
    If Not m_OriginalPats Is Nothing Then LoadPatients2 m_OriginalPats

End Sub

Private Sub lstPatienten_DblClick(ByVal blnCancel As MSForms.ReturnBoolean)
    
    If m_useDatabase Then
        Me.SetSelectedHospNum
    Else
        Me.SetSelectedBed
    End If
    
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
    
    strSort = objPat.AchterNaam & objPat.VoorNaam & objPat.PatientId
    strSort = IIf(m_onlyAdmitted, objPat.Bed & strSort, strSort)
    
    PatToSortString = strSort

End Function

Public Sub LoadPatients2(ByVal colPats As Collection)

    Dim objPat As ClassPatientDetails
    Dim arrSort() As Variant
    Dim varSort As Variant
    Dim strPat As String
    
    m_useDatabase = True
    
    Set m_Pats = New Collection
    Set m_OriginalPats = colPats
    
    Me.lstPatienten.Clear
    
    For Each objPat In colPats
        If m_onlyAdmitted Then
            If Not objPat.Bed = vbNullString And objPat.Afdeling = ModMetaVision.MetaVision_GetDepartment() Then
                ModArray.AddItemToVariantArray arrSort, PatToSortString(objPat)
            End If
        Else
            ModArray.AddItemToVariantArray arrSort, PatToSortString(objPat)
        End If
    Next objPat
    
    arrSort = SortArrayAtoZ(arrSort)
    
    For Each varSort In arrSort
        For Each objPat In colPats
            If PatToSortString(objPat) = varSort Then
                m_Pats.Add objPat
                strPat = objPat.ToString
                strPat = IIf(m_onlyAdmitted, objPat.Bed & " - " & strPat, strPat)
                
                Me.lstPatienten.AddItem strPat
            End If
        Next objPat
    Next varSort

End Sub

Public Sub SetSelectedHospNum()
    
    Dim objPat As ClassPatientDetails
    Dim strId As String

    If Me.lstPatienten.ListIndex > -1 Then
        Set objPat = m_Pats(Me.lstPatienten.ListIndex + 1)
        strId = objPat.PatientId
    End If
    
    ModBed.SetPatientHospitalNumber strId

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

Private Sub UserForm_Activate()

    CenterForm

End Sub

Private Sub UserForm_Terminate()
    
    If m_useDatabase Then
        Me.SetSelectedHospNum
    Else
        Me.SetSelectedBed
    End If

End Sub

Public Sub SetOnlyAdmittedTrue()

    m_onlyAdmitted = True
    chkAdmitted = m_onlyAdmitted

End Sub
