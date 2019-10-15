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

Private pLowerBound As Long

Public Property Get LowerBound() As Long
    LowerBound = pLowerBound
End Property

Public Property Let LowerBound(ByVal Value As Long)
    pLowerBound = Value
End Property


Private Sub Class_Initialize()
    ' Explicitly default to base 0
    pLowerBound = 0
End Sub

Public Function GetArray(Optional ByVal length As Long = 10, _
                         Optional ByVal ValueType As ValueTypes = ValueTypes.VariantVals, _
                         Optional ByVal ArrayType As ArrayTypes = ArrayTypes.OneDimension, _
                         Optional ByVal depth As Long) As Variant
    Dim localLength As Long
    Dim vals As IValuesList
       
    localLength = IIf(length <= 0, 1, length)
    
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
    GetArray = getArrayOfType(localLength, vals, ArrayType, depth)
End Function

Private Function getArrayOfType(ByVal length As Long, _
                                ByVal vals As IValuesList, _
                                ByVal ArrayType As ArrayTypes, _
                                Optional ByVal depth As Long) As Variant
    Dim result As Variant
    Select Case ArrayType
    Case ArrayTypes.Jagged
        result = getJaggedArray(length, vals, depth)
    Case ArrayTypes.MultiDimension
        result = getMultidimensionalArray(length, vals)
    Case ArrayTypes.OneDimension
        result = getOneDimensionalArray(length, vals)
    End Select
    getArrayOfType = result
End Function

Private Function getOneDimensionalArray(ByVal length As Long, ByVal vals As IValuesList) As Variant()
    Dim i As Long
    Dim arr() As Variant
    ReDim arr(pLowerBound To getUpperBound(length))
    For i = LBound(arr) To UBound(arr)
        If vals.IsObjectType Then
            Set arr(i) = vals.GetRandomValue()
        Else
            arr(i) = vals.GetRandomValue()
        End If
    Next
    getOneDimensionalArray = arr
End Function

Private Function getMultidimensionalArray(ByVal length As Long, ByVal vals As IValuesList) As Variant()
    Dim i As Long
    Dim j As Long
    Dim arr() As Variant
    ReDim arr(pLowerBound To getUpperBound(length), pLowerBound To getUpperBound(length))
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
            If vals.IsObjectType Then
                Set arr(i, j) = vals.GetRandomValue()
            Else
                arr(i, j) = vals.GetRandomValue()
            End If
        Next
    Next
    getMultidimensionalArray = arr
End Function

Private Function getJaggedArray(ByVal length As Long, _
                                ByVal vals As IValuesList, _
                                ByVal depth As Long, _
                                Optional ByVal currentDepth As Long) As Variant()
    Dim i As Long
    Dim localDepth As Long
    Dim localCurrentDepth As Long
    Dim result() As Variant
    
    'default to depth of 2
    localDepth = IIf(depth > 0, depth, 2)
    localCurrentDepth = currentDepth + 1
    ReDim result(pLowerBound To getUpperBound(length))
    For i = LBound(result) To UBound(result)
        If localCurrentDepth >= (localDepth - 1) Then
            result(i) = getOneDimensionalArray(length, vals)
        Else
            result(i) = getJaggedArray(length, vals, localDepth, localCurrentDepth)
        End If
    Next
    getJaggedArray = result
End Function

Public Function ConcatArraysOfSameStructure(ByVal ArrayType As ArrayTypes, ParamArray args() As Variant) As Variant()
    ' This method is exclusively for use as a test dependency and not for production use
    ' (same disclaimer applies to all methods in this class)
    Dim arg As Variant
    Dim result() As Variant
    Dim currentArray() As Variant
    Dim UpperBound As Long
    Dim currentIndex As Long
    
    UpperBound = -1
    
    If ArrayType = Unallocated Or ArrayType = Undefined Then
        'Handle Exception
    Else
        For Each arg In args
            If IsArray(arg) Then
                currentArray = arg
                UpperBound = UpperBound + GetArrayLength(currentArray)
            End If
        Next
        
        If ArrayType = MultiDimension Then
            ReDim result(UpperBound, UBound(currentArray, 2))
            For Each arg In args
                If IsArray(arg) Then
                    currentArray = arg
                    currentIndex = localMultiDimConcat(currentIndex, result, currentArray)
                End If
            Next
        Else
            ReDim result(UpperBound)
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

Public Function GetArrayLength(ByRef sourceArray() As Variant) As Long
    GetArrayLength = UBound(sourceArray) - LBound(sourceArray) + 1
End Function

Private Function getUpperBound(ByVal length As Long) As Long
    getUpperBound = length + pLowerBound - 1
End Function

