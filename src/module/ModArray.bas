Attribute VB_Name = "ModArray"
Option Explicit

Public Sub AddItemToVariantArray(arrItems() As Variant, varItem As Variant)

    Dim intU As Integer
    
    If Len(Join(arrItems)) = 0 Then
        ReDim arrItems(0)
    Else
        intU = UBound(arrItems) + 1
        ReDim Preserve arrItems(0 To intU)
    End If
    
    arrItems(intU) = varItem

End Sub


Public Sub AddItemToStringArray(arrItems() As String, strItem As String)

    Dim intU As Integer
    
    If Len(Join(arrItems)) = 0 Then
        ReDim arrItems(0)
    Else
        intU = UBound(arrItems) + 1
        ReDim Preserve arrItems(0 To intU)
    End If
    
    arrItems(intU) = strItem

End Sub


