VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Public Function getArray(Optional length As Long, Optional valType As ValueTypes, Optional arrType As ArrayTypes) As Variant
    ' remove 1 as arrays are base 0
    length = length - 1
    If length <= 0 Then length = 1
    If IsMissing(valType) Then valType = Variants
    If IsMissing(arrType) Then arrType = OneDimension
    Dim rtn As Variant
    Dim gen As IArrayGenerator
    Select Case valType
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
    getArray = getArrayOfType(length, gen, arrType)
End Function

Private Function getArrayOfType(length As Long, gen As IArrayGenerator, arrType As ArrayTypes) As Variant
    Dim rtn As Variant
    Select Case arrType
        Case ArrayTypes.jagged
            rtn = gen.getJaggedArray(length)
        Case ArrayTypes.MultiDimension
            rtn = gen.getMultidimensionalArray(length)
        Case ArrayTypes.OneDimension
            rtn = gen.getOneDimensionalArray(length)
    End Select
    getArrayOfType = rtn
End Function


