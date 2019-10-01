VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@Folder("Tests.Dependencies.ArrayGenerator")

Public Enum ValueTypes
    Booleans
    Bytes
    Doubles
    Longs
    Objects
    Strings
    Variants
End Enum

Public Enum ArrayTypes
    Undefined
    Unallocated
    OneDimension
    MultiDimension
    jagged
End Enum

Public Function getArray(Optional ByVal Length As Long, Optional ByVal ValueType As ValueTypes = ValueTypes.Variants, Optional ByVal ArrayType As ArrayTypes = ArrayTypes.OneDimension) As Variant
    Dim localLength As Long
    Dim gen As IArrayGenerator
       
    ' subtract 1 1 as arrays are base 0
    localLength = Length - 1
    If localLength <= 0 Then localLength = 1
    
    Select Case ValueType
    Case ValueTypes.Booleans
        Set gen = New ArrayGenerator_Booleans
    Case ValueTypes.Bytes
        Set gen = New ArrayGenerator_Bytes
    Case ValueTypes.Doubles
        Set gen = New ArrayGenerator_Doubles
    Case ValueTypes.Longs
        Set gen = New ArrayGenerator_Longs
    Case ValueTypes.Objects
        Set gen = New ArrayGenerator_Objects
    Case ValueTypes.Strings
        Set gen = New ArrayGenerator_Strings
    Case ValueTypes.Variants
        Set gen = New ArrayGenerator_Variants
    End Select
    getArray = getArrayOfType(localLength, gen, ArrayType)
End Function

Private Function getArrayOfType(ByVal Length As Long, ByVal gen As IArrayGenerator, ByVal ArrayType As ArrayTypes) As Variant
    Dim result As Variant
    Select Case ArrayType
    Case ArrayTypes.jagged
        result = gen.getJaggedArray(Length)
    Case ArrayTypes.MultiDimension
        result = gen.getMultidimensionalArray(Length)
    Case ArrayTypes.OneDimension
        result = gen.getOneDimensionalArray(Length)
    End Select
    getArrayOfType = result
End Function

