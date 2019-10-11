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

Public Function GetArray(Optional ByVal Length As Long = 10, _
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
    GetArray = getArrayOfType(localLength, vals, ArrayType)
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

Public Function ConcatArraysOfSameStructure(ByVal ArrayType As ArrayTypes, ParamArray args() As Variant) As Variant()
    ' This method is exclusively for use as a test dependency and not for production use
    ' (same disclaimer applies to all methods in this class)
    Dim arg As Variant
    Dim result() As Variant
    Dim currentArray() As Variant
    Dim upperBound As Long
    Dim currentIndex As Long
    
    upperBound = -1
    
    If ArrayType = Unallocated Or ArrayType = Undefined Then
        'Handle Exception
    Else
        For Each arg In args
            If IsArray(arg) Then
                currentArray = arg
                upperBound = upperBound + GetArrayLength(currentArray)
            End If
        Next
        
        If ArrayType = MultiDimension Then
            ReDim result(upperBound, UBound(currentArray, 2))
            For Each arg In args
                If IsArray(arg) Then
                    currentArray = arg
                    currentIndex = localMultiDimConcat(currentIndex, result, currentArray)
                End If
            Next
        Else
            ReDim result(upperBound)
            For Each arg In args
              If IsArray(arg) Then
                  currentArray = arg
                  currentIndex = localConcat(currentIndex, result, currentArray)
              End If
            Next
        End If
    End If
    
    ConcatArraysOfSameStructure = result

End Function

Private Function localMultiDimConcat(ByVal startingIndex As Long, ByRef result() As Variant, ByRef source() As Variant) As Long
    Dim i As Long
    Dim j As Long
    For i = LBound(source) To UBound(source)
        For j = LBound(source, 2) To UBound(source, 2)
            If IsObject(source(i, j)) Then
                Set result(i + startingIndex, j) = source(i, j)
            Else
                result(i + startingIndex, j) = source(i, j)
            End If
        Next
    Next
    localMultiDimConcat = startingIndex + i
End Function

Private Function localConcat(ByVal startingIndex As Long, ByRef result() As Variant, ByRef source() As Variant) As Long
    Dim i As Long
    For i = LBound(source) To UBound(source)
        If IsObject(source(i)) Then
            Set result(i + startingIndex) = source(i)
        Else
            result(i + startingIndex) = source(i)
        End If
    Next
    localConcat = startingIndex + i
End Function


Public Function GetArrayLength(ByRef SourceArray() As Variant) As Long
    GetArrayLength = UBound(SourceArray) - LBound(SourceArray) + 1
End Function
