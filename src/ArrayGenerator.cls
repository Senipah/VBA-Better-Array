VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Used to generate arrays of specified length and type. Dependency for 'TestModule_BetterArray'"
'@Folder("Tests.Dependencies.ArrayGenerator")
'@ModuleDescription("Used to generate arrays of specified length and type. Dependency for 'TestModule_BetterArray'")

Option Explicit

Public Enum ValueTypes
    AG_BOOLEAN
    AG_BYTE
    AG_DOUBLE
    AG_LONG
    AG_OBJECT
    AG_STRING
    AG_VARIANT
End Enum

Public Enum ArrayTypes
    AG_UNDEFINED
    AG_UNALLOCATED
    AG_ONEDIMENSION
    AG_MULTIDIMENSION
    AG_JAGGED
End Enum

Private pLowerBound As Long

Private Sub Class_Initialize()
    ' Explicitly default to base 0
    pLowerBound = 0
End Sub

'''''''''''''''''''''
' Public Properties '
'''''''''''''''''''''

Public Property Get LowerBound() As Long
    LowerBound = pLowerBound
End Property

Public Property Let LowerBound(ByVal Value As Long)
    pLowerBound = Value
End Property

'''''''''''''''''''
' Public Methods '
'''''''''''''''''''

Public Function GetArray( _
        Optional ByVal ValueType As ValueTypes = ValueTypes.AG_VARIANT, _
        Optional ByVal ArrayType As ArrayTypes = ArrayTypes.AG_ONEDIMENSION, _
        Optional ByVal Length As Long = 10, _
        Optional ByVal Depth As Long _
    ) As Variant
    Dim localLength As Long
    Dim Vals As IValuesList
       
    localLength = IIf(Length <= 0, 1, Length)
    
    Select Case ValueType
    Case ValueTypes.AG_BOOLEAN
        Set Vals = New ValuesList_Booleans
    Case ValueTypes.AG_BYTE
        Set Vals = New ValuesList_Bytes
    Case ValueTypes.AG_DOUBLE
        Set Vals = New ValuesList_Doubles
    Case ValueTypes.AG_LONG
        Set Vals = New ValuesList_Longs
    Case ValueTypes.AG_OBJECT
        Set Vals = New ValuesList_Objects
    Case ValueTypes.AG_STRING
        Set Vals = New ValuesList_Strings
    Case ValueTypes.AG_VARIANT
        Set Vals = New ValuesList_Variants
    End Select
    GetArray = getArrayOfType(localLength, Vals, ArrayType, Depth)
End Function

Public Function ConcatArraysOfSameStructure( _
        ByVal ArrayType As ArrayTypes, _
        ParamArray Args() As Variant _
    ) As Variant()
    ' This method is exclusively for use as a test dependency and not for production use
    ' (same disclaimer applies to all methods in this class)
    Dim arg As Variant
    Dim result() As Variant
    Dim currentArray() As Variant
    Dim UpperBound As Long
    Dim currentIndex As Long
    
    UpperBound = -1
    
    If ArrayType = AG_UNALLOCATED Or ArrayType = AG_UNDEFINED Then
        'Handle Exception
    Else
        For Each arg In Args
            If IsArray(arg) Then
                currentArray = arg
                UpperBound = UpperBound + getArrayLength(currentArray)
            End If
        Next
        
        If ArrayType = AG_MULTIDIMENSION Then
            ReDim result(UpperBound, UBound(currentArray, 2))
            For Each arg In Args
                If IsArray(arg) Then
                    currentArray = arg
                    currentIndex = localMultiDimConcat(currentIndex, result, currentArray)
                End If
            Next
        Else
            ReDim result(UpperBound)
            For Each arg In Args
              If IsArray(arg) Then
                  currentArray = arg
                  currentIndex = localConcat(currentIndex, result, currentArray)
              End If
            Next
        End If
    End If
    
    ConcatArraysOfSameStructure = result

End Function

Public Function getArrayLength(ByRef SourceArray() As Variant) As Long
    getArrayLength = UBound(SourceArray) - LBound(SourceArray) + 1
End Function

'''''''''''''''''''
' Private Methods '
'''''''''''''''''''

Private Function getArrayOfType( _
        ByVal Length As Long, _
        ByVal Vals As IValuesList, _
        ByVal ArrayType As ArrayTypes, _
        Optional ByVal Depth As Long _
    ) As Variant
    Dim result As Variant
    Select Case ArrayType
    Case ArrayTypes.AG_JAGGED
        result = getJaggedArray(Length, Vals, Depth)
    Case ArrayTypes.AG_MULTIDIMENSION
        result = getMultidimensionalArray(Length, Vals)
    Case ArrayTypes.AG_ONEDIMENSION
        result = getOneDimensionalArray(Length, Vals)
    End Select
    getArrayOfType = result
End Function

Private Function getOneDimensionalArray( _
        ByVal Length As Long, _
        ByVal Vals As IValuesList _
    ) As Variant()
    Dim i As Long
    Dim arr() As Variant
    ReDim arr(pLowerBound To getUpperBound(Length))
    For i = LBound(arr) To UBound(arr)
        If Vals.IsObjectType Then
            Set arr(i) = Vals.GetRandomValue()
        Else
            arr(i) = Vals.GetRandomValue()
        End If
    Next
    getOneDimensionalArray = arr
End Function

Private Function getMultidimensionalArray( _
        ByVal Length As Long, ByVal Vals As IValuesList) As Variant()
    Dim i As Long
    Dim j As Long
    Dim arr() As Variant
    ReDim arr(pLowerBound To getUpperBound(Length), pLowerBound To getUpperBound(Length))
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
            If Vals.IsObjectType Then
                Set arr(i, j) = Vals.GetRandomValue()
            Else
                arr(i, j) = Vals.GetRandomValue()
            End If
        Next
    Next
    getMultidimensionalArray = arr
End Function

Private Function getJaggedArray( _
        ByVal Length As Long, _
        ByVal Vals As IValuesList, _
        ByVal Depth As Long, _
        Optional ByVal CurrentDepth As Long _
    ) As Variant()
    Dim i As Long
    Dim localDepth As Long
    Dim localCurrentDepth As Long
    Dim result() As Variant
    
    'default to depth of 2
    localDepth = IIf(Depth > 0, Depth, 2)
    localCurrentDepth = CurrentDepth + 1
    ReDim result(pLowerBound To getUpperBound(Length))
    For i = LBound(result) To UBound(result)
        If localCurrentDepth >= (localDepth - 1) Then
            result(i) = getOneDimensionalArray(Length, Vals)
        Else
            result(i) = getJaggedArray(Length, Vals, localDepth, localCurrentDepth)
        End If
    Next
    getJaggedArray = result
End Function

Private Function localMultiDimConcat( _
        ByVal StartingIndex As Long, _
        ByRef result() As Variant, _
        ByRef Source() As Variant _
    ) As Long
    Dim i As Long
    Dim j As Long
    For i = LBound(Source) To UBound(Source)
        For j = LBound(Source, 2) To UBound(Source, 2)
            If IsObject(Source(i, j)) Then
                Set result(i + StartingIndex, j) = Source(i, j)
            Else
                result(i + StartingIndex, j) = Source(i, j)
            End If
        Next
    Next
    localMultiDimConcat = StartingIndex + i
End Function

Private Function localConcat( _
        ByVal StartingIndex As Long, _
        ByRef result() As Variant, _
        ByRef Source() As Variant _
    ) As Long
    Dim i As Long
    For i = LBound(Source) To UBound(Source)
        If IsObject(Source(i)) Then
            Set result(i + StartingIndex) = Source(i)
        Else
            result(i + StartingIndex) = Source(i)
        End If
    Next
    localConcat = StartingIndex + i
End Function

Private Function getUpperBound(ByVal Length As Long) As Long
    getUpperBound = Length + pLowerBound - 1
End Function













