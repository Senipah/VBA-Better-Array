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
    BooleanVals
    ByteVals
    DoubleVals
    LongVals
    ObjectVals
    StringVals
    VariantVals
End Enum

Public Enum ArrayTypes
    Undefined
    Unallocated
    OneDimension
    MultiDimension
    Jagged
End Enum

Public Function getArray(Optional ByVal Length As Long, _
                         Optional ByVal ValueType As ValueTypes = ValueTypes.VariantVals, _
                         Optional ByVal ArrayType As ArrayTypes = ArrayTypes.OneDimension) As Variant
    Dim localLength As Long
    Dim vals As IValuesList
       
    ' subtract 1 1 as arrays are base 0
    localLength = Length - 1
    If localLength <= 0 Then localLength = 1
    
    Select Case ValueType
    Case ValueTypes.BooleanVals
        Set vals = New ValuesList_Booleans
    Case ValueTypes.ByteVals
        Set vals = New ValuesList_Bytes
    Case ValueTypes.DoubleVals
        Set vals = New ValuesList_Doubles
    Case ValueTypes.LongVals
        Set vals = New ValuesList_Longs
    Case ValueTypes.ObjectVals
        Set vals = New ValuesList_Objects
    Case ValueTypes.StringVals
        Set vals = New ValuesList_Strings
    Case ValueTypes.VariantVals
        Set vals = New ValuesList_Variants
    End Select
    getArray = getArrayOfType(localLength, vals, ArrayType)
End Function

Private Function getArrayOfType(ByVal Length As Long, ByVal vals As IValuesList, ByVal ArrayType As ArrayTypes) As Variant
    Dim result As Variant
    Select Case ArrayType
    Case ArrayTypes.Jagged
        result = getJaggedArray(Length, vals)
    Case ArrayTypes.MultiDimension
        result = getMultidimensionalArray(Length, vals)
    Case ArrayTypes.OneDimension
        result = getOneDimensionalArray(Length, vals)
    End Select
    getArrayOfType = result
End Function

Private Function getOneDimensionalArray(ByVal Length As Long, ByVal vals As IValuesList) As Variant()
    Dim i As Long
    Dim arr() As Variant
    ReDim arr(Length)
    For i = 0 To Length
        If vals.IsObjectType Then
            Set arr(i) = vals.GetRandomValue()
        Else
            arr(i) = vals.GetRandomValue()
        End If
    Next
    getOneDimensionalArray = arr
End Function

Private Function getMultidimensionalArray(ByVal Length As Long, ByVal vals As IValuesList) As Variant()
    Dim i As Long
    Dim j As Long
    Dim arr() As Variant
    ReDim arr(Length, Length)
    For i = 0 To Length
        For j = 0 To Length
            If vals.IsObjectType Then
                Set arr(i, j) = vals.GetRandomValue()
            Else
                arr(i, j) = vals.GetRandomValue()
            End If
        Next
    Next
    getMultidimensionalArray = arr
End Function

Private Function getJaggedArray(ByVal Length As Long, ByVal vals As IValuesList) As Variant()
    Dim i As Long
    Dim j As Long
    Dim parentArray() As Variant
    Dim childArray() As Variant

    ReDim parentArray(Length)
    ReDim childArray(Length)
    For i = 0 To Length
        For j = 0 To Length
            If vals.IsObjectType Then
                Set childArray(j) = vals.GetRandomValue()
            Else
                childArray(j) = vals.GetRandomValue()
            End If
        Next
        parentArray(i) = childArray
    Next
    getJaggedArray = parentArray
End Function


