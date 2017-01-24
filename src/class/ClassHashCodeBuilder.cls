VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassHashCodeBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private lngConstant_ As Long
Private lngTotal_ As Long

Public Sub InitializeVariables(Optional lngInitialNonZeroOddNumber As Long, Optional lngMultiplierNonZeroOddNumber As Long)

    If (lngInitialNonZeroOddNumber = 0) Then
        lngConstant_ = 37
    ElseIf (lngInitialNonZeroOddNumber Mod 2 = 0) Then
        err.Raise Number:=5, Source:="Factory.createHashCodeBuilder", Description:="ClassHashCodeBuilder requires an odd initial value"
    Else
        lngConstant_ = lngMultiplierNonZeroOddNumber
    End If
    
    If (lngMultiplierNonZeroOddNumber = 0) Then
        'variable was not supplied
        lngTotal_ = 17
    ElseIf (lngMultiplierNonZeroOddNumber Mod 2 = 0) Then
        err.Raise Number:=5, Source:="Factory.createHashCodeBuilder", Description:="ClassHashCodeBuilder requires an odd multiplier"
    Else
        lngTotal_ = lngMultiplierNonZeroOddNumber
    End If
    
End Sub

'' *****************************
'' ****** PRIMITIVE TYPES ******
'' *****************************

Public Function AppendLong(lngValue As Long) As ClassHashCodeBuilder

    lngTotal_ = lngTotal_ * lngConstant_ + lngValue
    Set AppendLong = Me

End Function

Public Function AppendInteger(intValue As Integer) As ClassHashCodeBuilder

    lngTotal_ = lngTotal_ * lngConstant_ + intValue
    Set AppendInteger = Me

End Function

Public Function AppendByte(bytValue As Byte) As ClassHashCodeBuilder
    
    lngTotal_ = lngTotal_ * lngConstant_ + bytValue
    Set AppendByte = Me

End Function

Public Function AppendSingle(sngValue As Single) As ClassHashCodeBuilder
    
    lngTotal_ = lngTotal_ * lngConstant_ + sngValue
    Set AppendSingle = Me

End Function

Public Function AppendDouble(dblValue As Double) As ClassHashCodeBuilder
    
    lngTotal_ = lngTotal_ * lngConstant_ + dblValue
    Set AppendDouble = Me

End Function

Public Function AppendBoolean(blnValue As Boolean) As ClassHashCodeBuilder
    
    If blnValue = True Then
        lngTotal_ = lngTotal_ * lngConstant_ + blnValue
    End If
    
    Set AppendBoolean = Me

End Function

'' ***************************************
'' ****** ARRAYS OF PRIMITIVE TYPES ******
'' ***************************************

Public Function AppendLongArray(lngArr() As Long) As ClassHashCodeBuilder

    Dim lngVal As Long
    Dim varVal As Variant
    
    If UBound(lngArr, 1) - LBound(lngArr, 1) + 1 = 0 Then
        lngTotal_ = lngTotal_ * lngConstant_
    Else
        For Each varVal In lngArr
            AppendLong CLng(varVal)
        Next varVal
    End If
    
    Set AppendLongArray = Me

End Function

Public Function AppendIntegerArray(intArr() As Integer) As ClassHashCodeBuilder
    
    Dim varVal As Variant
    
    
    If UBound(intArr, 1) - LBound(intArr, 1) + 1 = 0 Then
        lngTotal_ = lngTotal_ * lngConstant_
    Else
        For Each varVal In intArr
            AppendInteger CInt(varVal)
        Next varVal
    End If
        
    Set AppendIntegerArray = Me

End Function

Public Function AppendByteArray(bytArr() As Byte) As ClassHashCodeBuilder
    
    Dim varVal As Variant
    
    If UBound(bytArr, 1) - LBound(bytArr, 1) + 1 = 0 Then
        lngTotal_ = lngTotal_ * lngConstant_
    Else
        For Each varVal In bytArr
            AppendByte CByte(varVal)
        Next varVal
    End If
        
    Set AppendByteArray = Me

End Function

Public Function AppendSingleArray(sngArr() As Single) As ClassHashCodeBuilder
     
    Dim varVal As Variant
      
    If UBound(sngArr, 1) - LBound(sngArr, 1) + 1 = 0 Then
        lngTotal_ = lngTotal_ * lngConstant_
    Else
        For Each varVal In sngArr
            AppendSingle CSng(varVal)
        Next varVal
    End If
        
    Set AppendSingleArray = Me

End Function

Public Function AppendDoubleArray(dblArr() As Double) As ClassHashCodeBuilder
     
    Dim varVal As Variant
     
    If UBound(dblArr, 1) - LBound(dblArr, 1) + 1 = 0 Then
        lngTotal_ = lngTotal_ * lngConstant_
    Else
        For Each varVal In dblArr
            AppendDouble CDbl(varVal)
        Next varVal
    End If
       
    Set AppendDoubleArray = Me

End Function

Public Function AppendBooleanArray(blnArr() As Boolean) As ClassHashCodeBuilder
     
    Dim varVal As Variant
    
    If UBound(blnArr, 1) - LBound(blnArr, 1) + 1 = 0 Then
        lngTotal_ = lngTotal_ * lngConstant_
    Else
        For Each varVal In blnArr
            AppendBoolean CBool(varVal)
        Next varVal
    End If
    
    Set AppendBooleanArray = Me

End Function

Public Function HashCode() As Long
    
    HashCode = lngTotal_

End Function


