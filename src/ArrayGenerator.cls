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
'@Folder("VBABetterArray.Tests.Dependencies.ArrayGenerator")
'@ModuleDescription("Used to generate arrays of specified length and type. Dependency for 'TestModule_BetterArray'")
'@IgnoreModule ProcedureNotUsed

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

Public Enum AG_ArrayTypes
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
        Optional ByVal ArrayType As ArrayTypes = AG_ArrayTypes.AG_ONEDIMENSION, _
        Optional ByVal Length As Long = 10, _
        Optional ByVal Depth As Long, _
        Optional ByVal ArrayWidth As Long = 10 _
    ) As Variant
    Dim LocalLength As Long
    Dim Vals As IValuesList
       
    LocalLength = IIf(Length <= 0, 1, Length)
    
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
    GetArray = GetArrayOfType(LocalLength, Vals, ArrayType, Depth, ArrayWidth)
End Function

Public Function ConcatArraysOfSameStructure( _
        ByVal ArrayType As ArrayTypes, _
        ParamArray Args() As Variant _
    ) As Variant()
    ' This method is exclusively for use as a test dependency and not for production use
    ' (same disclaimer applies to all methods in this class)
    Dim Arg As Variant
    Dim Result() As Variant
    Dim CurrentArray() As Variant
    Dim UpperBound As Long
    Dim CurrentIndex As Long
    
    UpperBound = -1
    
    '@Ignore EmptyIfBlock
    If ArrayType = AG_UNALLOCATED Or ArrayType = AG_UNDEFINED Then
        'Handle Exception
    Else
        For Each Arg In Args
            If IsArray(Arg) Then
                CurrentArray = Arg
                UpperBound = UpperBound + GetArrayLength(CurrentArray)
            End If
        Next
        
        If ArrayType = AG_MULTIDIMENSION Then
            ReDim Result(UpperBound, UBound(CurrentArray, 2))
            For Each Arg In Args
                If IsArray(Arg) Then
                    CurrentArray = Arg
                    CurrentIndex = LocalMultiDimConcat(CurrentIndex, Result, CurrentArray)
                End If
            Next
        Else
            ReDim Result(UpperBound)
            For Each Arg In Args
              If IsArray(Arg) Then
                  CurrentArray = Arg
                  CurrentIndex = LocalConcat(CurrentIndex, Result, CurrentArray)
              End If
            Next
        End If
    End If
    
    ConcatArraysOfSameStructure = Result

End Function

Public Function GetArrayLength(ByRef SourceArray() As Variant) As Long
    GetArrayLength = UBound(SourceArray) - LBound(SourceArray) + 1
End Function

'''''''''''''''''''
' Private Methods '
'''''''''''''''''''

Private Function GetArrayOfType( _
        ByVal Length As Long, _
        ByVal Vals As IValuesList, _
        ByVal ArrayType As ArrayTypes, _
        ByVal Depth As Long, _
        ByVal ArrayWidth As Long _
    ) As Variant
    Dim Result As Variant
    Select Case ArrayType
    Case AG_ArrayTypes.AG_JAGGED
        Result = GetJaggedArray(Length, Vals, Depth, ArrayWidth)
    Case AG_ArrayTypes.AG_MULTIDIMENSION
        Result = GetMultidimensionalArray(Length, Vals, ArrayWidth)
    Case AG_ArrayTypes.AG_ONEDIMENSION
        Result = GetOneDimensionalArray(Length, Vals)
    End Select
    GetArrayOfType = Result
End Function

Private Function GetOneDimensionalArray( _
        ByVal Length As Long, _
        ByVal Vals As IValuesList _
    ) As Variant()
    Dim i As Long
    Dim Arr() As Variant
    ReDim Arr(pLowerBound To GetUpperBound(Length))
    For i = LBound(Arr) To UBound(Arr)
        If Vals.IsObjectType Then
            Set Arr(i) = Vals.GetRandomValue
        Else
            Arr(i) = Vals.GetRandomValue
        End If
    Next
    GetOneDimensionalArray = Arr
End Function

Private Function GetMultidimensionalArray( _
        ByVal Length As Long, _
        ByVal Vals As IValuesList, _
        ByVal ArrayWidth As Long _
    ) As Variant()
    Dim i As Long
    Dim j As Long
    Dim Arr() As Variant
    ReDim Arr(pLowerBound To GetUpperBound(Length), pLowerBound To GetUpperBound(ArrayWidth))
    For i = LBound(Arr, 1) To UBound(Arr, 1)
        For j = LBound(Arr, 2) To UBound(Arr, 2)
            If Vals.IsObjectType Then
                Set Arr(i, j) = Vals.GetRandomValue
            Else
                Arr(i, j) = Vals.GetRandomValue
            End If
        Next
    Next
    GetMultidimensionalArray = Arr
End Function

Private Function GetJaggedArray( _
        ByVal Length As Long, _
        ByVal Vals As IValuesList, _
        ByVal Depth As Long, _
        ByVal ArrayWidth As Long, _
        Optional ByVal CurrentDepth As Long _
    ) As Variant()
    Dim i As Long
    Dim LocalDepth As Long
    Dim LocalCurrentDepth As Long
    Dim Result() As Variant
    
    'default to depth of 2
    LocalDepth = IIf(Depth > 0, Depth, 2)
    LocalCurrentDepth = CurrentDepth + 1
    ReDim Result(pLowerBound To GetUpperBound(Length))
    For i = LBound(Result) To UBound(Result)
        If LocalCurrentDepth >= (LocalDepth - 1) Then
            Result(i) = GetOneDimensionalArray(ArrayWidth, Vals)
        Else
            Result(i) = GetJaggedArray(Length, Vals, LocalDepth, ArrayWidth, LocalCurrentDepth)
        End If
    Next
    GetJaggedArray = Result
End Function

Private Function LocalMultiDimConcat( _
        ByVal StartingIndex As Long, _
        ByRef Result() As Variant, _
        ByRef Source() As Variant _
    ) As Long
    Dim i As Long
    Dim j As Long
    For i = LBound(Source) To UBound(Source)
        For j = LBound(Source, 2) To UBound(Source, 2)
            If IsObject(Source(i, j)) Then
                Set Result(i + StartingIndex, j) = Source(i, j)
            Else
                Result(i + StartingIndex, j) = Source(i, j)
            End If
        Next
    Next
    LocalMultiDimConcat = StartingIndex + i
End Function

Private Function LocalConcat( _
        ByVal StartingIndex As Long, _
        ByRef Result() As Variant, _
        ByRef Source() As Variant _
    ) As Long
    Dim i As Long
    For i = LBound(Source) To UBound(Source)
        If IsObject(Source(i)) Then
            Set Result(i + StartingIndex) = Source(i)
        Else
            Result(i + StartingIndex) = Source(i)
        End If
    Next
    LocalConcat = StartingIndex + i
End Function

Private Function GetUpperBound(ByVal Length As Long) As Long
    GetUpperBound = Length + pLowerBound - 1
End Function


