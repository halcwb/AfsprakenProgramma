Attribute VB_Name = "ModCollection"
Option Explicit
Option Base 0

'Returns True if the Collection has the key, varKey. Otherwise, returns False
Public Function CollectionHasKey(varKey As Variant, objCol As Collection) As Boolean
    
    On Error GoTo ColHasKeyErr
    
    objCol varKey
    CollectionHasKey = True
    
    Exit Function

ColHasKeyErr:

    CollectionHasKey = False

End Function

'Returns True if the Collection CollectionContains an element equal to varValue
Public Function CollectionContains(varValue As Variant, objCol As Collection) As Boolean

    CollectionContains = (CollectionIndexOf(varValue, objCol) >= 0)

End Function


'Returns the first lngIndex of an element equal to varValue. If the Collection
'does not contain such an element, returns -1.
Public Function CollectionIndexOf(varValue As Variant, objCol As Collection) As Long

    Dim lngIndex As Long
    
    For lngIndex = 1 To objCol.Count Step 1
        If objCol(lngIndex) = varValue Then
            CollectionIndexOf = lngIndex
            Exit Function
        End If
    Next lngIndex
    
    CollectionIndexOf = -1
    
End Function


'Sorts the given collection using the Arrays.MergeSort algorithm.
' O(n log(n)) time
' O(n) space
Public Sub CollectionSort(objCol As Collection, Optional objC As IVariantComparator)

    Dim varA() As Variant
    
    If objCol.Count = 0 Then Exit Sub
    
    varA = CollectionToArray(objCol)
    ModArray.ArraySort varA(), objC
    
    Set objCol = CollectionFromArray(varA())

End Sub

'Returns an array which exactly matches this collection.
' Note: This function is not safe for concurrent modification.
Public Function CollectionToArray(objCol As Collection) As Variant

    Dim varA() As Variant
    Dim lngN As Long
    
    ReDim varA(0 To objCol.Count)
    
    For lngN = 0 To objCol.Count - 1
        varA(lngN) = objCol(lngN + 1)
    Next lngN
    
    CollectionToArray = varA()

End Function

'Returns a Collection which exactly matches the given Array
' Note: This function is not safe for concurrent modification.
Public Function CollectionFromArray(a() As Variant) As Collection

    Dim objCol As Collection
    Dim varElement As Variant
    
    Set objCol = New Collection
    For Each varElement In a
        objCol.Add varElement
    Next varElement
    
    Set CollectionFromArray = objCol
    
End Function

'Adds all elements from the source collection, colSrc, to the destination collection, colDest.
'Returns true if the destination collection changed as a result of this operation; false otherwise.
Public Function CollectionAddAllFromCol(colSrc As Collection, colDest As Collection) As Boolean

    Dim lngCount As Long
    Dim varElement As Variant
    
    lngCount = colDest.Count
    
    For Each varElement In colSrc
        colDest.Add varElement
    Next varElement
    
    CollectionAddAllFromCol = (colDest.Count = lngCount)
    
End Function

'Adds all elements from the source array, varSrc, to the destination collection, colDest
'Returns true if the destination collection changed as a result of this operation; false otherwise.
Public Function CollectionAddAllFromArray(varSrc() As Variant, colDest As Collection) As Boolean
    
    Dim lngCount As Long
    Dim varElement As Variant
    
    lngCount = colDest.Count
    
    For Each varElement In varSrc
        colDest.Add varElement
    Next varElement
    
    CollectionAddAllFromArray = (colDest.Count = lngCount)
    
End Function

Public Sub CollectionAddDistinctStringNotEmpty(objColl As Collection, ByVal varValue As Variant)

    If Not IsEmpty(varValue) Then
        varValue = CStr(varValue)
        If Not StringIsZeroOrEmpty(varValue) And Not CollectionContains(varValue, objColl) Then
            objColl.Add varValue
        End If
    End If

End Sub

Public Function ConcatenateCollection(objColl As Collection, ByVal strDel As String) As String

    If objColl.Count > 0 Then
        ConcatenateCollection = Join(CollectionToArray(objColl), strDel)
    Else
        ConcatenateCollection = vbNullString
    End If

End Function
