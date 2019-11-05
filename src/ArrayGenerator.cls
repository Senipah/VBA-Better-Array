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

Public Property Let LowerBound(ByVal value As Long)
    pLowerBound = value
End Property

'''''''''''''''''''
' Public Methods '
'''''''''''''''''''

Public Function GetArray( _
        Optional ByVal valueType As ValueTypes = ValueTypes.AG_VARIANT, _
        Optional ByVal ArrayType As ArrayTypes = ArrayTypes.AG_ONEDIMENSION, _
        Optional ByVal Length As Long = 10, _
        Optional ByVal depth As Long _
    ) As Variant
    Dim localLength As Long
    Dim vals As IValuesList
       
    localLength = IIf(Length <= 0, 1, Length)
    
    Select Case valueType
    Case ValueTypes.AG_BOOLEAN
        Set vals = New ValuesList_Booleans
    Case ValueTypes.AG_BYTE
        Set vals = New ValuesList_Bytes
    Case ValueTypes.AG_DOUBLE
        Set vals = New ValuesList_Doubles
    Case ValueTypes.AG_LONG
        Set vals = New ValuesList_Longs
    Case ValueTypes.AG_OBJECT
        Set vals = New ValuesList_Objects
    Case ValueTypes.AG_STRING
        Set vals = New ValuesList_Strings
    Case ValueTypes.AG_VARIANT
        Set vals = New ValuesList_Variants
    End Select
    GetArray = GetArrayOfType(localLength, vals, ArrayType, depth)
End Function

Public Function ConcatArraysOfSameStructure( _
        ByVal ArrayType As ArrayTypes, _
        ParamArray args() As Variant _
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
        For Each arg In args
            If IsArray(arg) Then
                currentArray = arg
                UpperBound = UpperBound + GetArrayLength(currentArray)
            End If
        Next
        
        If ArrayType = AG_MULTIDIMENSION Then
            ReDim result(UpperBound, UBound(currentArray, 2))
            For Each arg In args
                If IsArray(arg) Then
                    currentArray = arg
                    currentIndex = LocalMultiDimConcat(currentIndex, result, currentArray)
                End If
            Next
        Else
            ReDim result(UpperBound)
            For Each arg In args
              If IsArray(arg) Then
                  currentArray = arg
                  currentIndex = LocalConcat(currentIndex, result, currentArray)
              End If
            Next
        End If
    End If
    
    ConcatArraysOfSameStructure = result

End Function

Public Function GetArrayLength(ByRef sourceArray() As Variant) As Long
    GetArrayLength = UBound(sourceArray) - LBound(sourceArray) + 1
End Function

'''''''''''''''''''
' Private Methods '
'''''''''''''''''''

Private Function GetArrayOfType( _
        ByVal Length As Long, _
        ByVal vals As IValuesList, _
        ByVal ArrayType As ArrayTypes, _
        Optional ByVal depth As Long _
    ) As Variant
    Dim result As Variant
    Select Case ArrayType
    Case ArrayTypes.AG_JAGGED
        result = GetJaggedArray(Length, vals, depth)
    Case ArrayTypes.AG_MULTIDIMENSION
        result = GetMultidimensionalArray(Length, vals)
    Case ArrayTypes.AG_ONEDIMENSION
        result = GetOneDimensionalArray(Length, vals)
    End Select
    GetArrayOfType = result
End Function

Private Function GetOneDimensionalArray( _
        ByVal Length As Long, _
        ByVal vals As IValuesList _
    ) As Variant()
    Dim i As Long
    Dim arr() As Variant
    ReDim arr(pLowerBound To GetUpperBound(Length))
    For i = LBound(arr) To UBound(arr)
        If vals.IsObjectType Then
            Set arr(i) = vals.GetRandomValue()
        Else
            arr(i) = vals.GetRandomValue()
        End If
    Next
    GetOneDimensionalArray = arr
End Function

Private Function GetMultidimensionalArray( _
        ByVal Length As Long, _
        ByVal vals As IValuesList _
    ) As Variant()
    Dim i As Long
    Dim j As Long
    Dim arr() As Variant
    ReDim arr(pLowerBound To GetUpperBound(Length), pLowerBound To GetUpperBound(Length))
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
            If vals.IsObjectType Then
                Set arr(i, j) = vals.GetRandomValue()
            Else
                arr(i, j) = vals.GetRandomValue()
            End If
        Next
    Next
    GetMultidimensionalArray = arr
End Function

Private Function GetJaggedArray( _
        ByVal Length As Long, _
        ByVal vals As IValuesList, _
        ByVal depth As Long, _
        Optional ByVal currentDepth As Long _
    ) As Variant()
    Dim i As Long
    Dim localDepth As Long
    Dim localCurrentDepth As Long
    Dim result() As Variant
    
    'default to depth of 2
    localDepth = IIf(depth > 0, depth, 2)
    localCurrentDepth = currentDepth + 1
    ReDim result(pLowerBound To GetUpperBound(Length))
    For i = LBound(result) To UBound(result)
        If localCurrentDepth >= (localDepth - 1) Then
            result(i) = GetOneDimensionalArray(Length, vals)
        Else
            result(i) = GetJaggedArray(Length, vals, localDepth, localCurrentDepth)
        End If
    Next
    GetJaggedArray = result
End Function

Private Function LocalMultiDimConcat( _
        ByVal startingIndex As Long, _
        ByRef result() As Variant, _
        ByRef source() As Variant _
    ) As Long
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
    LocalMultiDimConcat = startingIndex + i
End Function

Private Function LocalConcat( _
        ByVal startingIndex As Long, _
        ByRef result() As Variant, _
        ByRef source() As Variant _
    ) As Long
    Dim i As Long
    For i = LBound(source) To UBound(source)
        If IsObject(source(i)) Then
            Set result(i + startingIndex) = source(i)
        Else
            result(i + startingIndex) = source(i)
        End If
    Next
    LocalConcat = startingIndex + i
End Function

Private Function GetUpperBound(ByVal Length As Long) As Long
    GetUpperBound = Length + pLowerBound - 1
End Function















'v0.0.9
