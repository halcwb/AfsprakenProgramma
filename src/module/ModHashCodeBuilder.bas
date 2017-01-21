Attribute VB_Name = "ModHashCodeBuilder"
Option Explicit

Public Function NewHashCodeBuilder(Optional lngInitialNonZeroOddNumber As Long, Optional lngMultiplierNonZeroOddNumber As Long) As ClassHashCodeBuilder

    Set NewHashCodeBuilder = New ClassHashCodeBuilder
    NewHashCodeBuilder.InitializeVariables lngInitialNonZeroOddNumber, lngMultiplierNonZeroOddNumber
    
End Function

