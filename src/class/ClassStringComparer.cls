VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassStringComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IVariantComparator

Public Function IVariantComparator_Compare(varA As Variant, varB As Variant) As Long

    IVariantComparator_Compare = StrComp(varA, varB, vbTextCompare)

End Function

