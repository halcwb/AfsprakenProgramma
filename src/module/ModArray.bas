Attribute VB_Name = "ModArray"
Option Compare Text
Option Explicit

Private Const INSERTIONSORT_THRESHOLD As Long = 7

Public Sub AddItemToVariantArray(ByRef arrItems() As Variant, ByVal varItem As Variant)

    Dim intU As Integer
    
    If Len(Join(arrItems)) = 0 Then
        ReDim arrItems(0)
    Else
        intU = UBound(arrItems) + 1
        ReDim Preserve arrItems(0 To intU)
    End If
    
    arrItems(intU) = varItem

End Sub

Public Sub AddItemToStringArray(ByRef arrItems() As String, ByVal strItem As String)

    Dim intU As Integer
    
    If Len(Join(arrItems)) = 0 Then
        ReDim arrItems(0)
    Else
        intU = UBound(arrItems) + 1
        ReDim Preserve arrItems(0 To intU)
    End If
    
    arrItems(intU) = strItem

End Sub

'Sorts the array using the MergeSort algorithm (follows the Java legacyMergesort algorithm
'O(n*log(n)) time; O(n) space
Public Sub ArraySort(ByRef varA() As Variant, Optional ByRef objC As IVariantComparator)

    If objC Is Nothing Then
        MergeSort CreateCopyOfVariantArray(varA), varA, 0, VariantArrayLength(varA), 0, objC ' ToDo Factory.newNumericComparator
    Else
        MergeSort CreateCopyOfVariantArray(varA), varA, 0, VariantArrayLength(varA), 0, objC
    End If
    
End Sub

Private Sub MergeSort(ByRef varSrc() As Variant, ByRef varDest() As Variant, lngLow As Long, lngHigh As Long, lngOff As Long, ByRef objC As IVariantComparator)

    Dim lngLength As Long
    Dim lngDestLow As Long
    Dim lngDestHigh As Long
    Dim lngMid As Long
    
    Dim lngN As Long
    Dim lngP As Long
    Dim lngQ As Long
    Dim lngJ As Long
    
    lngLength = lngHigh - lngLow
    
    ' insertion sort on small arrays
    If lngLength < INSERTIONSORT_THRESHOLD Then
        lngN = lngLow
        Do While lngN < lngHigh
            lngJ = lngN
            Do While True
                If (lngJ <= lngLow) Then
                    Exit Do
                End If
                If (objC.Compare(varDest(lngJ - 1), varDest(lngJ)) <= 0) Then
                    Exit Do
                End If
                Swap varDest, lngJ, lngJ - 1
                lngJ = lngJ - 1                  'decrement lngJ
            Loop
            lngN = lngN + 1                      'increment lngN
        Loop
        Exit Sub
    End If
    
    'recursively sort halves of varDest into varSrc
    lngDestLow = lngLow
    lngDestHigh = lngHigh
    lngLow = lngLow + lngOff
    lngHigh = lngHigh + lngOff
    lngMid = (lngLow + lngHigh) / 2
    MergeSort varDest, varSrc, lngLow, lngMid, -lngOff, objC
    MergeSort varDest, varSrc, lngMid, lngHigh, -lngOff, objC
    
    'if list is already sorted, we're done
    If objC.Compare(varSrc(lngMid - 1), varSrc(lngMid)) <= 0 Then
        ArrayCopyVariants varSrc, lngLow, varDest, lngDestLow, lngLength - 1
        Exit Sub
    End If
    
    'merge sorted halves into varDest
    lngN = lngDestLow
    lngP = lngLow
    lngQ = lngMid
    Do While lngN < lngDestHigh
        If (lngQ >= lngHigh) Then
            varDest(lngN) = varSrc(lngP)
            lngP = lngP + 1
        Else
            'Otherwise, check if p<mid AND varSrc(lngP) preceeds scr(lngQ)
            'See description of following idom at: http://stackoverflow.com/a/3245183/3795219
            Select Case True
            Case lngP >= lngMid, objC.Compare(varSrc(lngP), varSrc(lngQ)) > 0
                varDest(lngN) = varSrc(lngQ)
                lngQ = lngQ + 1
            Case Else
                varDest(lngN) = varSrc(lngP)
                lngP = lngP + 1
            End Select
        End If
       
        lngN = lngN + 1
    Loop
    
End Sub

'Sorts the array using the MergeSort algorithm (follows the Java legacyMergesort algorithm
'O(n*log(n)) time; O(n) space
Public Sub SortObjects(ByRef varA() As Object, ByRef objC As IObjectComparator)

    If objC Is Nothing Then
        err.Raise 3, "Arrays.sortObjects", "No IObjectComparator Provided to the sortObjects method."
    End If
    
    MergeSortObjects CreateCopyOfObjecttArray(varA), varA, 0, ObjectArrayLength(varA), 0, objC
    
End Sub

Private Sub MergeSortObjects(ByRef objSrc() As Object, ByRef objDest() As Object, lngLow As Long, lngHigh As Long, lngOff As Long, ByRef objC As IObjectComparator)

    Dim lngLenght As Long
    Dim lngDestLow As Long
    Dim lngDestHigh As Long
    Dim lngMid As Long
    
    Dim lngN As Long
    Dim lngP As Long
    Dim lngQ As Long
    Dim lngJ As Long
    
    lngLenght = lngHigh - lngLow
    
    ' insertion sort on small arrays
    If lngLenght < INSERTIONSORT_THRESHOLD Then
        lngN = lngLow
        Do While lngN < lngHigh
            lngJ = lngN
            Do While True
                If (lngJ <= lngLow) Then
                    Exit Do
                End If
                If (objC.Compare(objDest(lngJ - 1), objDest(lngJ)) <= 0) Then
                    Exit Do
                End If
                SwapObjects objDest, lngJ, lngJ - 1
                lngJ = lngJ - 1                  'decrement lngJ
            Loop
            lngN = lngN + 1                      'increment lngN
        Loop
        Exit Sub
    End If
    
    'recursively sort halves of objDest into objSrc
    lngDestLow = lngLow
    lngDestHigh = lngHigh
    lngLow = lngLow + lngOff
    lngHigh = lngHigh + lngOff
    lngMid = (lngLow + lngHigh) / 2
    MergeSortObjects objDest, objSrc, lngLow, lngMid, -lngOff, objC
    MergeSortObjects objDest, objSrc, lngMid, lngHigh, -lngOff, objC
    
    'if list is already sorted, we're done
    If objC.Compare(objSrc(lngMid - 1), objSrc(lngMid)) <= 0 Then
        ArrayCopyObjects objSrc, lngLow, objDest, lngDestLow, lngLenght - 1
        Exit Sub
    End If
    
    'merge sorted halves into objDest
    lngN = lngDestLow
    lngP = lngLow
    lngQ = lngMid
    Do While lngN < lngDestHigh
        If (lngQ >= lngHigh) Then
            objDest(lngN) = objSrc(lngP)
            lngP = lngP + 1
        Else
            'Otherwise, check if p<mid AND objSrc(lngP) preceeds scr(lngQ)
            'See description of following idom at: http://stackoverflow.com/a/3245183/3795219
            Select Case True
            Case lngP >= lngMid, objC.Compare(objSrc(lngP), objSrc(lngQ)) > 0
                objDest(lngN) = objSrc(lngQ)
                lngQ = lngQ + 1
            Case Else
                objDest(lngN) = objSrc(lngP)
                lngP = lngP + 1
            End Select
        End If
       
        lngN = lngN + 1
    Loop
    
End Sub

Private Sub Swap(varArr() As Variant, lngA As Long, lngB As Long)
    
    Dim varT As Variant
    
    varT = varArr(lngA)
    varArr(lngA) = varArr(lngB)
    varArr(lngB) = varT

End Sub

Private Sub SwapObjects(objArr() As Object, lngA As Long, lngB As Long)
    
    Dim objT As Object
    
    objT = objArr(lngA)
    objArr(lngA) = objArr(lngB)
    objArr(lngB) = objT

End Sub

Public Function CreateCopyOfVariantArray(ByRef varOriginal() As Variant) As Variant()
    
    Dim varDest() As Variant
    
    varDest = Array()
    ReDim varDest(LBound(varOriginal) To UBound(varOriginal))
    CopyRangeVariants varOriginal, LBound(varOriginal), UBound(varOriginal), varDest
    
    CreateCopyOfVariantArray = varDest

End Function

Private Sub CopyRangeVariants(varSource() As Variant, lngBegin As Long, lngEnd As Long, varDest() As Variant)
    
    Dim lngK As Long
    
    For lngK = lngBegin To lngEnd Step 1
        varDest(lngK) = varSource(lngK)
    Next lngK

End Sub

Private Sub CopyRangeObjects(objSource() As Object, lngBegin As Long, lngEnd As Long, objDest() As Object)
    
    Dim lngK As Long
    
    For lngK = lngBegin To lngEnd Step 1
        objDest(lngK) = objSource(lngK)
    Next lngK

End Sub

Public Function CreateCopyOfObjecttArray(ByRef objOriginal() As Object) As Object()

    Dim objDest() As Object
    
    objDest = Array()
    ReDim objDest(LBound(objOriginal) To UBound(objOriginal))
    CopyRangeObjects objOriginal, LBound(objOriginal), UBound(objOriginal), objDest
    
    CreateCopyOfObjecttArray = objDest

End Function

'Copies an array from the specified source array, beginning at the specified position, to the specified position in the destination array
Public Sub ArrayCopyVariants(ByRef varSrc() As Variant, lngSrcPos As Long, ByRef varDst() As Variant, lngDstPos As Long, lngLength As Long)
    
    Dim intN As Long
    
    'Check if all offsets and lengths are non negative
    If lngSrcPos < 0 Or lngDstPos < 0 Or lngLength < 0 Then
        err.Raise 9, , "negative value supplied"
    End If
     
    'Check if ranges are valid
    If lngLength + lngSrcPos > UBound(varSrc) Then
        err.Raise 9, , "Not enough elements to ArrayCopyVariants, src+length: " & lngSrcPos + lngLength & ", UBound(varSrc): " & UBound(varSrc)
    End If
    If lngLength + lngDstPos > UBound(varDst) Then
        err.Raise 9, , "Not enough room in destination array. dstPos+length: " & lngDstPos + lngLength & ", UBound(varDst): " & UBound(varDst)
    End If
    
    intN = 0
    'ArrayCopyVariants varSrc elements to varDst
    Do While lngLength > intN
        varDst(lngDstPos + intN) = varSrc(lngSrcPos + intN)
        intN = intN + 1
    Loop
    
End Sub

'Copies an array from the specified source array, beginning at the specified position, to the specified position in the destination array
Public Sub ArrayCopyObjects(ByRef objSrc() As Object, lngSrcPos As Long, ByRef objDst() As Object, lngDstPos As Long, lngLength As Long)
    
    Dim intN As Long
    
    'Check if all offsets and lengths are non negative
    If lngSrcPos < 0 Or lngDstPos < 0 Or lngLength < 0 Then
        err.Raise 9, , "negative value supplied"
    End If
     
    'Check if ranges are valid
    If lngLength + lngSrcPos > UBound(objSrc) Then
        err.Raise 9, , "Not enough elements to ArrayCopyVariants, src+length: " & lngSrcPos + lngLength & ", UBound(objSrc): " & UBound(objSrc)
    End If
    If lngLength + lngDstPos > UBound(objDst) Then
        err.Raise 9, , "Not enough room in destination array. dstPos+length: " & lngDstPos + lngLength & ", UBound(objDst): " & UBound(objDst)
    End If
    
    intN = 0
    'ArrayCopyVariants objSrc elements to objDst
    Do While lngLength > intN
        objDst(lngDstPos + intN) = objSrc(lngSrcPos + intN)
        intN = intN + 1
    Loop
    
End Sub


'Adds all elements from the source collection, colSrc, to the destination collection, varDest.
'Returns true if the destination collection changed as a result of this operation; false otherwise.
Public Function ArrayAddAllFromCol(ByRef colSrc As Collection, ByRef varDest() As Variant) As Boolean

    Dim lngCount As Long
    Dim lngN As Long
    Dim varElement As Variant
    
    lngCount = VariantArrayLength(varDest)
    lngN = 1
    ReDim Preserve varDest(lngCount + colSrc.Count)
    
    For Each varElement In colSrc
        Set varDest(lngCount + lngN) = varElement
    Next varElement
    
    ArrayAddAllFromCol = (VariantArrayLength(varDest) = lngCount)
    
End Function

'Adds all elements from the source array, varSrc, to the destination collection, colDest
'Returns true if the destination collection changed as a result of this operation; false otherwise.
Public Function AddAllFromArray(ByRef varSrc() As Variant, ByRef colDest As Collection) As Boolean

    Dim lngCount As Long
    Dim lngN As Long
    Dim varElement As Variant
    
    lngCount = colDest.Count
    lngN = 1
    
    For Each varElement In varSrc
        Set colDest(lngCount + lngN) = varElement
    Next varElement
    
    AddAllFromArray = (colDest.Count = lngCount)
    
End Function

Public Function VariantArrayLength(ByRef varA() As Variant) As Long

    VariantArrayLength = UBound(varA) - LBound(varA) + 1
    
End Function

Public Function ObjectArrayLength(ByRef objA() As Object) As Long

    ObjectArrayLength = UBound(objA) - LBound(objA) + 1
    
End Function

Public Function ArrayToString(ByRef varA() As Variant) As String

    Dim varElement As Variant
    
    If VariantArrayLength(varA) <= 0 Then
        ArrayToString = "[]"
    ElseIf VariantArrayLength(varA) = 1 Then
        ArrayToString = "[ " & varA(UBound(varA)) & " ]"
    Else
        ArrayToString = "[ "
        For Each varElement In varA
            ArrayToString = ArrayToString & varElement & " "
        Next varElement
        ArrayToString = ArrayToString & " ]"
    End If
    
End Function

Public Function StringArrayItem(ByRef strArr() As String, ByVal intItem As Integer) As String

    Dim strResult As String
    
    If UBound(strArr) < intItem Then
        strResult = vbNullString
    Else
        strResult = strArr(intItem)
    End If
    
    StringArrayItem = strResult

End Function


Private Sub TestStringArrayItem()

    Dim strArr() As String
    
    strArr = Split("test !", " ")
    MsgBox StringArrayItem(strArr, 3)

End Sub
