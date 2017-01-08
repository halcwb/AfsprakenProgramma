Attribute VB_Name = "ModList"
Option Explicit

' === Region with common Pick List code


Public Function GetSelectedListCount(ByRef lstList As MSForms.ListBox) As Integer

    Dim intN As Integer
    Dim intC As Integer
    
    intC = 0
    For intN = 0 To lstList.ListCount - 1
        If lstList.Selected(intN) Then intC = intC + 1
    Next intN
    
    GetSelectedListCount = intC

End Function

Public Sub LoadListItems(ByRef lstList As MSForms.ListBox, ByRef colItems As Collection)

    Dim varItem As Variant
    
    For Each varItem In colItems
        lstList.AddItem varItem
    Next varItem

End Sub

Public Sub SelectListItem(ByRef lstList As MSForms.ListBox, ByVal intN As Integer)

    lstList.Selected(intN - 2) = True

End Sub

Public Function IsListItemSelected(ByRef lstList As MSForms.ListBox, ByVal intN As Integer) As Boolean

    IsListItemSelected = lstList.Selected(intN - 2)

End Function

Public Sub UnselectListItem(ByRef lstList As MSForms.ListBox, ByVal intN As Integer)

    lstList.Selected(intN - 2) = False

End Sub

Public Function HasSelectedListItems(ByRef lstList As MSForms.ListBox) As Boolean
    
    HasSelectedListItems = Not GetFirstSelectedListItem(lstList, False) = 1

End Function

Public Function GetFirstSelectedListItem(ByRef lstList As MSForms.ListBox, ByVal blnUnSelect As Boolean) As Integer

    Dim intN As Integer
    Dim intC As Integer
    
    intC = lstList.ListCount - 1
    For intN = 0 To intC
        If lstList.Selected(intN) Then
            If blnUnSelect Then lstList.Selected(intN) = False
            GetFirstSelectedListItem = intN + 2
            Exit Function
        End If
    Next intN
    
    GetFirstSelectedListItem = 1

End Function
