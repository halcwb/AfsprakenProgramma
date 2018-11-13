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

Private m_pats As Collection
Private m_onlyAdmitted As Boolean
Private m_useDatabase As Boolean

Private Sub chkAdmitted_Click()

    m_onlyAdmitted = chkAdmitted.Value
    LoadPatients2 m_pats

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
    Set m_pats = colPats
    
    For Each objPat In colPats
        Me.lstPatienten.AddItem objPat.ToString
    Next objPat

End Sub

Function SortArrayAtoZ(myArray As Variant)

    Dim lngI As Long
    Dim lngJ As Long
    Dim varTemp
    
    'Sort the Array A-Z
    For lngI = LBound(myArray) To UBound(myArray) - 1
        For lngJ = lngI + 1 To UBound(myArray)
            If UCase(myArray(lngI)) > UCase(myArray(lngJ)) Then
                varTemp = myArray(lngJ)
                myArray(lngJ) = myArray(lngI)
                myArray(lngI) = varTemp
            End If
        Next lngJ
    Next lngI
    
    SortArrayAtoZ = myArray

End Function

Public Sub LoadPatients2(ByVal colPats As Collection)

    Dim objPat As ClassPatientDetails
    Dim arrLN() As Variant
    Dim strLN As Variant
    
    m_useDatabase = True
    
    Set m_pats = New Collection
    Me.lstPatienten.Clear
    
    For Each objPat In colPats
        If m_onlyAdmitted Then
            If Not objPat.Bed = vbNullString Then ModArray.AddItemToVariantArray arrLN, objPat.AchterNaam
        Else
            ModArray.AddItemToVariantArray arrLN, objPat.AchterNaam
        End If
    Next objPat
    
    arrLN = SortArrayAtoZ(arrLN)
    
    For Each strLN In arrLN
        For Each objPat In colPats
            If objPat.AchterNaam = strLN Then
                m_pats.Add objPat
                Me.lstPatienten.AddItem objPat.ToString
            End If
        Next objPat
    Next

End Sub

Public Sub SetSelectedHospNum()
    
    Dim objPat As ClassPatientDetails
    Dim strId As String

    If Me.lstPatienten.ListIndex > -1 Then
        Set objPat = m_pats(Me.lstPatienten.ListIndex + 1)
        strId = objPat.PatientId
    End If
    
    ModBed.SetPatientHospitalNumber strId

End Sub

Public Sub SetSelectedBed()
    
    Dim objPat As ClassPatientInfo
    Dim strBed As String

    If Me.lstPatienten.ListIndex > -1 Then
        Set objPat = m_pats(Me.lstPatienten.ListIndex + 1)
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
