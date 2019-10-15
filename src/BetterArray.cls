VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BetterArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@Folder("BetterArray")

'''''''''''''''''''
' Constant Fields '
'''''''''''''''''''

Private Const DEFAULT_CAPACITY As Long = 4
Private Const MAX_ARRAY_LENGTH = &H7FEFFFFF
Private Enum ArrayTypes
    Undefined
    Unallocated
    OneDimension
    MultiDimension
    Jagged
End Enum

Private Enum ErrorCodes
    StartCode = vbObjectError + 512
    ExpectedRangeObject
    MaxDimensionsLimit
    EndCode
End Enum

Private Type ErrorDefinition
    number As Long
    source As String
    description As String
End Type


''''''''''
' Fields '
''''''''''

Private pCapacity As Long
Private pLength As Long
Private pLowerBound As Long
Private pItems() As Variant
Private pArrayType As ArrayTypes
Private ErrorDefinitions(ErrorCodes.StartCode To ErrorCodes.EndCode) As ErrorDefinition


'''''''''''''''
' Constructor '
'''''''''''''''

Private Sub Class_Initialize()
    Me.Capacity = DEFAULT_CAPACITY
    PopulateErrorDefinitions
End Sub

''''''''''''''
' Destructor '
''''''''''''''
'@Ignore EmptyMethod
Private Sub Class_Terminate()

End Sub

'''''''''''''''''''''
' Public Properties '
'''''''''''''''''''''

Public Property Get Capacity() As Long
    If IsEmpty(pCapacity) Then Me.Capacity = DEFAULT_CAPACITY
    Capacity = pCapacity
End Property

Public Property Let Capacity(ByVal Value As Long)
    If Value < 0 Then Err.Raise 9
    If Value <> pCapacity Then
        If pCapacity > 0 Then
            If (UBound(pItems) - LBound(pItems) + 1) <> Value Then
                Dim newItems() As Variant
                newItems = pItems
                ReDim Preserve newItems(pLowerBound To (Value + pLowerBound - 1))
                pItems = newItems
            End If
        Else
            ReDim pItems(pLowerBound To (DEFAULT_CAPACITY + pLowerBound - 1))
        End If
        pCapacity = UBound(pItems) - pLowerBound + 1
    End If
End Property

Public Property Get Length() As Long
    Length = pLength
End Property

Public Property Get UpperBound()
    'UpperBound = IIf(pLength = 0, pLength, pLength + pLowerBound - 1)
    UpperBound = pLength + pLowerBound - 1
End Property

Public Property Get LowerBound() As Long
    LowerBound = pLowerBound
End Property

Public Property Let LowerBound(ByVal Value As Long)
    If Value <> pLowerBound And Value >= 0 Then
        pLowerBound = Value
        Rebase
    End If
End Property

' For releases ensure to set the DefaultMember annotation & attribute.
' Disable on dev branch as it breaks code inspections:
' https://github.com/rubberduck-vba/Rubberduck/issues/5222
Public Property Get Item(ByVal Index As Long)
    If Index <= pLength Then
        If IsObject(pItems(Index)) Then
            Set Item = pItems(Index)
        Else
            Item = pItems(Index)
        End If
    Else
        Err.Raise 9
    End If
End Property

Public Property Let Item(ByVal Index As Long, ByVal element As Variant)
    If Me.UpperBound >= Index Then
        If Index < Me.LowerBound Then
            Me.Unshift element
        Else
            If IsObject(element) Then
                Set pItems(Index) = element
            Else
                pItems(Index) = element
            End If
        End If
    Else
        Me.Push element
    End If
End Property

Public Property Get Items() As Variant
    Dim result() As Variant
    result = getInternalItems()
    If pArrayType = ArrayTypes.MultiDimension And isJaggedArray(result) Then result = jaggedToMulti(result)
    Items = result
End Property

Public Property Let Items(ByVal Values As Variant)
    Dim localLowerBound As Long
    Dim localValues() As Variant
    If IsArray(Values) Then
        localValues = Values
        pArrayType = getArrayType(localValues)
        If pArrayType = ArrayTypes.Undefined _
           Or pArrayType = ArrayTypes.Unallocated Then
            ' TOD: Define custom error
            Err.Raise 9
        Else
            If pArrayType = ArrayTypes.MultiDimension Then
                localValues = multiToJagged(localValues)
            End If
            localLowerBound = LBound(localValues)
            pItems = localValues
            pLength = UBound(pItems) - LBound(pItems) + 1
            If localLowerBound <> pLowerBound Then
                Rebase
            End If
            Me.Capacity = pLength
        End If
    Else
        ' If Values is not an array then clear our internal array and push the Values arg
        ' TODO: Document and write tests
        Me.Clear.Push Values
    End If

End Property

''''''''''''''''''''''
' Private Properties '
''''''''''''''''''''''

''''''''''''''''''
' Public Methods '
''''''''''''''''''

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
Public Function Push(ParamArray args() As Variant) As Variant
    Dim element As Variant

    If pArrayType = ArrayTypes.Unallocated Or _
       pArrayType = ArrayTypes.Undefined Then
        pArrayType = ArrayTypes.OneDimension
    End If

    For Each element In args
        If IsArray(element) And pArrayType = ArrayTypes.OneDimension Then
            pArrayType = ArrayTypes.Jagged
        End If
        If pLength = pCapacity Then
            EnsureCapacity pLength + 1
        End If
        If IsObject(element) Then
            Set pItems(pLength) = element
        Else
            pItems(pLength) = element
        End If
        pLength = pLength + 1
    Next
    Push = pLength
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
Public Function Pop() As Variant
    Dim newItems() As Variant
    Dim result As Variant

    If pLength > 0 Then
        result = pItems(Me.UpperBound)
        pLength = pLength - 1
        ReDim newItems(pLowerBound To Me.UpperBound)
        Me.Items = copyRange(newItems)
    End If
    Pop = result
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
Public Function Shift() As Variant
    Dim result As Variant
    Dim newItems() As Variant
    If pLength > 0 Then
        result = pItems(pLowerBound)
        ReDim newItems((pLowerBound + 1) To Me.UpperBound)
        Me.Items = copyRange(newItems)
    End If
    Shift = result
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
Public Function Unshift(ParamArray args() As Variant) As Long
    Dim newItems() As Variant
    Dim oldItems() As Variant
    Dim oldType As ArrayTypes
    
    newItems = args
    oldType = pArrayType
    oldItems = getInternalItems()
    Me.Items = newItems
    If oldType <> Unallocated And oldType <> Undefined Then
        Me.Concat oldItems
        pArrayType = oldType
    End If
    Unshift = Me.Length
End Function

' TODO: Add ParseFromString method - compliment from ToString
'@Ignore FunctionReturnValueNotUsed
Public Function ParseFromString(ByVal SourceString As String, _
                                Optional ByVal ValueSeparator As String = ",", _
                                Optional ByVal ArrayOpenDelimiter As String, _
                                Optional ByVal ArrayClosingDelimiter As String) As BetterArray
    Dim opener As String
    Dim closer As String
    Dim arrayCount As Long
    Dim ArraysAreDelimited As Boolean
    Dim result() As Variant
    Dim tempStringArray() As String
    
    If Len(SourceString) > 0 Then
        If ArrayOpenDelimiter = vbNullString And ArrayClosingDelimiter = vbNullString Then
            ' no array delimiters supplied
            ' see if the diff in ascii values are for first and last chars are less than 3
            ' this covers () [] {} <>
            Dim firstChar As String
            Dim lastChar As String
            firstChar = Left$(SourceString, 1)
            lastChar = Right$(SourceString, 1)
            If Asc(lastChar) - Asc(firstChar) < 3 Then
                opener = firstChar
                closer = lastChar
                ArraysAreDelimited = True
            End If
        Else
            ' array delimiters supplied
            opener = ArrayOpenDelimiter
            closer = ArrayClosingDelimiter
            ArraysAreDelimited = True
        End If
        
        If ArraysAreDelimited Then
            arrayCount = Len(SourceString) - Len(Replace(SourceString, opener, vbNullString))
            If arrayCount = 1 Then
                tempStringArray = Split(Mid$(SourceString, 2, Len(SourceString) - 2), ValueSeparator)
                result = duckTypeStringArray(tempStringArray)
            ElseIf arrayCount > 1 Then
                'nested array
                result = extractNestedArrays(SourceString, ValueSeparator, opener, closer)
            Else
                ' handle arrayCount of 0
                ' NOTE: is this possible??
            End If
        Else
            tempStringArray = Split(SourceString, ValueSeparator)
            result = duckTypeStringArray(tempStringArray)
        End If
    Else
        ' nullstring passed
        'handle exception
    End If
    Me.Items = result
    Set ParseFromString = Me
End Function

Private Function extractNestedArrays(ByVal SourceString As String, _
                              ByVal ValueSeparator As String, _
                              ByVal opener As String, _
                              ByVal closer As String, _
                              Optional ByRef cursor As Long = 2) As Variant()
    Dim localResult() As Variant
    Dim i As Long
    
    ' default cursor to 2 because we know the first character will be an array delimiter
    
    Do While cursor < Len(SourceString) And cursor <> 0
        If Mid$(SourceString, cursor + 1, 1) = opener Then
            If i > 0 Then
                ReDim Preserve localResult(i)
            Else
                ReDim localResult(i)
            End If
            cursor = cursor + 1
            localResult(i) = extractNestedArrays( _
                    SourceString, _
                    ValueSeparator, _
                    opener, _
                    closer, _
                    cursor _
                )
            i = i + 1
        Else
            Do Until cursor = 0
                If i > 0 Then
                    ReDim Preserve localResult(i)
                Else
                    ReDim localResult(i)
                End If
                localResult(i) = extractArraysAtSameDimension( _
                    SourceString, _
                    ValueSeparator, _
                    opener, _
                    closer, _
                    cursor _
                )
                i = i + 1
                If cursor < 0 Then
                    cursor = Abs(cursor)
                    Exit Do
                End If
            Loop
        End If
    Loop
    
    
    extractNestedArrays = localResult
End Function

Private Function extractArraysAtSameDimension(ByVal SourceString As String, _
                                              ByVal ValueSeparator As String, _
                                              ByVal opener As String, _
                                              ByVal closer As String, _
                                              ByRef cursor As Long) As Variant()
    Dim localResult() As Variant
    Dim segment() As String
    Dim segmentEnd As Long
    
    ' NOTE: didn't want to use regex so as to aviod external dependencies
    
    Do While Mid$(SourceString, cursor + 1, 1) = opener And cursor < Len(SourceString)
        cursor = cursor + 1
    Loop
    
    segmentEnd = cursor
    Do While Mid$(SourceString, segmentEnd + 1, 1) <> closer And segmentEnd <= Len(SourceString)
        segmentEnd = segmentEnd + 1
    Loop
    
    segment = Split(Mid$(SourceString, cursor + 1, segmentEnd - cursor), ValueSeparator)
    localResult = duckTypeStringArray(segment)
    
    cursor = segmentEnd + 1
    Do While Mid$(SourceString, cursor + 1, 1) <> opener And cursor <= Len(SourceString)
        cursor = cursor + 1
        If Mid$(SourceString, cursor, 1) = closer Then
            'invert cursor if last array at current dimension
            cursor = cursor * -1
            Exit Do
        End If
    Loop
    
    If cursor >= Len(SourceString) Then cursor = 0
    
    extractArraysAtSameDimension = localResult
End Function

Private Function duckTypeStringArray(ByRef arr() As String) As Variant()
    Dim i As Long
    Dim localLowerBound As Long
    Dim localUpperBound As Long
    Dim result() As Variant
    Dim element As String
    
    localLowerBound = LBound(arr)
    localUpperBound = UBound(arr)
    ReDim result(localLowerBound To localUpperBound)
    For i = localLowerBound To localUpperBound
        element = arr(i)
        If UCase$(element) = "TRUE" Or UCase$(element) = "FALSE" Then
            result(i) = CBool(element)
        ElseIf IsNumeric(element) Then
            result(i) = Val(element)
        Else
            result(i) = Trim$(element)
        End If
    Next
    duckTypeStringArray = result
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
Public Function ToString(Optional ByVal PrettyPrint As Boolean) As String
    Dim ArrayType As ArrayTypes
    Dim result As String
    Dim localItems() As Variant
    
    localItems = getInternalItems()
    ArrayType = getArrayType(localItems)
    If Not ArrayType = ArrayTypes.Undefined And _
       Not ArrayType = ArrayTypes.Unallocated Then
        If ArrayType = ArrayTypes.MultiDimension Then
            localItems = multiToJagged(localItems)
        End If
        result = recursiveToString(localItems, PrettyPrint)
    End If
    ToString = result
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
Public Function Includes(ByVal SearchElement As Variant, Optional ByVal FromIndex As Long) As Boolean
    Dim localLength As Long
    Dim currentIndex As Long
    Dim arr() As Variant
    
    arr = getInternalItems()
    localLength = Me.Length
    
    If localLength = 0 Then
        Includes = False
        Exit Function
    End If
        
    If FromIndex > Me.LowerBound Then
        currentIndex = FromIndex
    ElseIf FromIndex = 0 Then
        'assume 0 val should always be from beginning - cover cases where val is missing
        currentIndex = Me.LowerBound
    Else
        currentIndex = localLength + FromIndex
    End If
    
    Do While currentIndex < localLength
        If arr(currentIndex) = SearchElement Then
            Includes = True
            Exit Function
        End If
        currentIndex = currentIndex + 1
    Loop
    Includes = False
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
Public Function Keys() As Variant()
    Dim i As Long
    Dim localLowerBound As Long
    Dim localUpperBound As Long
    Dim result() As Variant
    
    localLowerBound = Me.LowerBound
    localUpperBound = Me.UpperBound
    ReDim result(localLowerBound To localUpperBound)
    For i = localLowerBound To localUpperBound: result(i) = i: Next
    Keys = result
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
Public Function Max(ParamArray args() As Variant) As Variant
    Dim i As Long
    Dim arr() As Variant
    Dim result As Variant
    
    If UBound(args) < LBound(args) Then          'Test if no args provided
        arr = getInternalItems()
    Else
        arr = args
        If IsArray(arr) Then
            ' exit function if array is multi dimensional
            If isMultidimensionalArray(arr) Then
                Exit Function
            End If
        End If
    End If
    
    ' extract first array if argument(s) are arrays
    Do While IsArray(arr(0))
        arr = arr(0)
    Loop
    
    For i = LBound(arr) To UBound(arr)
        If Not IsArray(arr(i)) And Not IsObject(arr(i)) Then
            If IsEmpty(result) Then
                result = arr(i)
            Else
                If arr(i) > result Then result = arr(i)
            End If
        End If
    Next
    
    Max = result
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
Public Function Min(ParamArray args() As Variant) As Variant
    Dim i As Long
    Dim arr() As Variant
    Dim result As Variant
    
    If UBound(args) < LBound(args) Then          'Test if no args provided
        arr = getInternalItems()
    Else
        arr = args
        If IsArray(arr) Then
            ' exit function if array is multi dimensional
            If isMultidimensionalArray(arr) Then
                Exit Function
            End If
        End If
    End If
    
    ' extract first array if argument(s) are arrays
    Do While IsArray(arr(0))
        arr = arr(0)
    Loop
        
    For i = LBound(arr) To UBound(arr)
        If Not IsArray(arr(i)) And Not IsObject(arr(i)) Then
            If IsEmpty(result) Then
                result = arr(i)
            Else
                If arr(i) > result Then result = arr(i)
            End If
        End If
    Next
    
    Min = result
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
Public Function Slice(ByVal StartIndex As Long, Optional ByVal EndIndex As Long) As Variant()
    '@Ignore ShadowedDeclaration
    Dim localLength As Long
    Dim relativeStart As Long
    Dim relativeEnd As Long
    Dim oldIndex As Long
    Dim final As Long
    Dim count As Long
    Dim newIndex As Long
    Dim arr() As Variant
    Dim result() As Variant
    
    arr = getInternalItems()
    localLength = Me.Length
    relativeStart = StartIndex
    If relativeStart < 0 Then
        oldIndex = Max((localLength + relativeStart), 0)
    Else
        oldIndex = Min(relativeStart, localLength)
    End If
    If EndIndex = 0 Then
        relativeEnd = localLength
    Else
        relativeEnd = EndIndex
    End If
    If relativeEnd < 0 Then
        final = Max((localLength + relativeEnd), 0)
    Else
        final = Min(relativeEnd, localLength)
    End If
    
    count = Max(final - oldIndex, 0)
    ReDim result(newIndex To count - 1)
    Do While oldIndex < final
        If oldIndex >= LBound(arr) And oldIndex <= UBound(arr) Then
            result = arr(oldIndex)
            newIndex = newIndex + 1
            oldIndex = oldIndex + 1
        End If
    Loop
    
    Slice = result
End Function

' TODO: Implement, Document & Test ToRange method
'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
Public Function ToExcelRange(ByRef destination As Object, Optional ByVal TransposeValues As Boolean) As Object
    Dim localRange As Object
    Dim localItems() As Variant
    Dim depth As Long
    Dim lengthRows As Long
    Dim lengthColumns As Long
    
    lengthRows = 1 'default
    lengthColumns = 1 'default
    
    If TypeName(destination) = "Range" Then
        localItems = getInternalItems()
        depth = getJaggedArrayDepth(localItems)
        If depth > 0 Then 'check array is allocated
            If depth = 1 Then
                '1d array
                localItems = ensureScalar1DArray(localItems)
                If TransposeValues Then
                    lengthRows = Me.Length
                    localItems = Transpose_1DArray(localItems)
                Else
                    lengthColumns = Me.Length
                End If
            ElseIf depth > 1 Then
                'jagged array - multidim arrays should also be stored as jagged internally
                Const outputDepth As Long = 2 'need 2 dimensions for tabular representation output
                localItems = jaggedToMulti(localItems, outputDepth, EnsureScalar:=True)
                If TransposeValues Then
                    lengthRows = UBound(localItems, 2) - LBound(localItems, 2) + 1
                    lengthColumns = UBound(localItems, 1) - LBound(localItems, 1) + 1
                    localItems = Transpose_2DArray(localItems)
                Else
                    lengthRows = UBound(localItems, 1) - LBound(localItems, 1) + 1
                    lengthColumns = UBound(localItems, 2) - LBound(localItems, 2) + 1
                End If
            End If
            Set localRange = destination.Resize( _
                RowSize:=lengthRows, _
                ColumnSize:=lengthColumns _
            )
            localRange.Value = localItems
        Else
            'array must be unalloc
            'do nothing(?)
        End If
    Else
        RaiseError ErrorCodes.ExpectedRangeObject, "ToExcelRange()", "destination"
    End If
    
    Set ToExcelRange = localRange
End Function

''''''''''''''''''''''''''''
' Chainable Public Methods '
''''''''''''''''''''''''''''

' TODO: Document & Test - Clone
'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
Public Function Clone() As BetterArray
    Dim result As BetterArray
    Set result = New BetterArray
    result.LowerBound = Me.LowerBound
    result.Items = Me.Items
    Set Clone = result
End Function

' TODO: Document & Test - ResetToDefault
'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
Public Function ResetToDefault() As BetterArray
    pLowerBound = 0
    pItems = getEmptyArray()
    pLength = 0
    Me.Capacity = DEFAULT_CAPACITY
    Set ResetToDefault = Me
End Function

' TODO: Document & Test - ResetToDefault
'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
Public Function Clear() As BetterArray
    pItems = getEmptyArray()
    pLength = 0
    Me.Capacity = pCapacity
    Set Clear = Me
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
Public Function Concat(ParamArray args() As Variant) As BetterArray
    Dim arg As Variant
    Dim element As Variant
    Dim localArray() As Variant
    Dim localType As ArrayTypes
    
    
    For Each arg In args
        If IsArray(arg) Then
            localArray = arg
            localType = getArrayType(localArray)
            Select Case pArrayType
                Case ArrayTypes.Unallocated, ArrayTypes.Undefined, ArrayTypes.OneDimension
                    pArrayType = localType
            End Select
            If localType = MultiDimension Then
                localArray = multiToJagged(localArray)
            End If
            For Each element In localArray
                Me.Push element
            Next
        Else
            Me.Push arg
        End If
    Next
    Set Concat = Me
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
Public Function CopyFromCollection(ByVal SourceCollection As Collection) As BetterArray
    If SourceCollection Is Nothing Then Err.Raise 448
    Dim i As Long
    Dim newItems() As Variant
    pLength = SourceCollection.count
    If pLength = 0 Then
        newItems = getEmptyArray
    Else
        ReDim newItems(pLowerBound To (pLength - pLowerBound - 1))
        For i = 1 To pLength
            newItems(i + pLowerBound - 1) = SourceCollection.Item(i)
        Next
    End If
    Me.Items = newItems
    Set CopyFromCollection = Me
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
Public Function Sort(Optional ByVal SortColumn As Long) As BetterArray
    Dim localItems() As Variant
    Dim sortedItems() As Variant
    Dim ArrayType As ArrayTypes
    
    localItems = getInternalItems()
    ArrayType = getArrayType(localItems)
    If ArrayType = ArrayTypes.Unallocated Or _
       ArrayType = ArrayTypes.Undefined Then
        sortedItems = getEmptyArray
    Else
        If Me.Length > 0 Then
            If ArrayType <> OneDimension Then
                If ArrayType = ArrayTypes.MultiDimension Then
                    localItems = multiToJagged(localItems)
                End If
                Dim depth As Long
                depth = getJaggedArrayDepth(localItems)
                If depth > 2 Then
                    Debug.Print "Cannot sort greater than 2 dimensions"
                    Exit Function
                End If
            End If
            quickSort localItems, LBound(localItems), UBound(localItems), ArrayType, SortColumn
        End If
        sortedItems = localItems
    End If
    Me.Items = sortedItems
    Set Sort = Me
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
Public Function CopyWithin(ByVal Target As Long, _
                           Optional ByVal StartIndex As Long, _
                           Optional ByVal EndIndex As Long) As BetterArray
    Dim localLength As Long
    Dim relativeTarget As Long
    Dim relativeStart As Long
    Dim relativeEnd As Long
    Dim idx As Long
    Dim from As Long
    Dim final As Long
    Dim count As Long
    Dim direction As Long
    Dim arr() As Variant
    
    arr = getInternalItems()
    localLength = Me.Length
    relativeTarget = Target
    
    If relativeTarget < 0 Then
        idx = Max((localLength + relativeTarget), 0)
    Else
        idx = Min(relativeTarget, localLength)
    End If
    
    relativeStart = StartIndex
    
    If relativeStart < 0 Then
        from = Max((localLength + relativeStart), 0)
    Else
        from = Min(relativeStart, localLength)
    End If
    
    If EndIndex = 0 Then
        relativeEnd = localLength
    Else
        relativeEnd = EndIndex
    End If
    
    If relativeEnd < 0 Then
        final = Max((localLength + relativeEnd), 0)
    Else
        final = Min(relativeEnd, localLength)
    End If
    
    count = Min(final - from, localLength - idx)
    
    If from < idx And idx < from + count Then
        direction = -1
        from = from + count - 1
        idx = idx + count - 1
    Else
        direction = 1
    End If
    
    Do While count > 0
        If from >= LBound(arr) And from <= UBound(arr) Then
            arr(idx) = arr(from)
        End If
        from = from + direction
        idx = idx + direction
        count = count - 1
    Loop
    
    Me.Items = arr
    Set CopyWithin = Me
End Function

' TODO: add support for 2d arrays.
'@Ignore ShadowedDeclaration
'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
Public Function Filter(ByVal Match As Variant, Optional ByVal Exclude As Boolean) As BetterArray
    Dim element As Variant
    Dim oldArr() As Variant
    Dim newArr() As Variant
    Dim i As Long
    
    oldArr = getInternalItems()
    ReDim newArr(LBound(oldArr), UBound(oldArr))
    i = Me.LowerBound
    For Each element In oldArr
        If (Exclude And element = Match) Or (Not Exclude And element <> Match) Then
            newArr(i) = element
            i = i + 1
        End If
    Next
    ReDim Preserve newArr(Me.LowerBound, i)
    Me.Items = newArr
    Set Filter = Me
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
Public Function Reverse() As BetterArray
    Dim i As Long
    Dim localUpperBound As Long
    Dim arr() As Variant
    Dim result() As Variant
    arr = getInternalItems()
    localUpperBound = Me.UpperBound
    ReDim result(Me.LowerBound To localUpperBound)
    For i = localUpperBound To Me.LowerBound Step -1
        result = arr(i)
    Next
    Me.Items = result
    Set Reverse = Me
End Function



' TODO: Implement, Document & Test Shuffle method

' TODO: Implement, Document & Test Flatten method

' TODO: add algos from C++ STL https://en.cppreference.com/w/cpp/algorithm

' TODO: Implement remove method to delete entry by index

'''''''''''''''''''
' Private Methods '
'''''''''''''''''''

Private Sub PopulateErrorDefinitions()
    Dim source As String
    source = TypeName(Me)
    ErrorDefinitions(ErrorCodes.ExpectedRangeObject) = ErrorDefinitionFactory( _
        number:=ErrorCodes.ExpectedRangeObject, _
        source:=source, _
        description:="Range Object Expected" _
    )
    ErrorDefinitions(ErrorCodes.MaxDimensionsLimit) = ErrorDefinitionFactory( _
        number:=ErrorCodes.MaxDimensionsLimit, _
        source:=source, _
        description:="Cannot convert arrays with more than 20 dimensions to a multidimensional array!" _
    )
End Sub

Private Function ErrorDefinitionFactory(ByVal number As Long, ByVal source As String, ByVal description As String) As ErrorDefinition
    Dim result As ErrorDefinition
    result.number = number
    result.source = source
    result.description = description
    ErrorDefinitionFactory = result
End Function

Private Sub RaiseError(ByVal ErrorCode As ErrorCodes, ByVal Caller As String, Optional ByVal ArgName As String)
    Dim currentError As ErrorDefinition
    Dim localArgName As String
    If ArgName <> vbNullString Then localArgName = "Variable " & ArgName & ": "
    currentError = ErrorDefinitions(ErrorCode)
    Err.Raise currentError.number, _
              currentError.source & "." & Caller, _
              localArgName & currentError.description
End Sub

'Used by ToExcelRange
Private Function ensureScalar1DArray(ByRef arr() As Variant) As Variant()
    Dim i As Long
    Dim localLowerBound As Long
    Dim localUpperBound As Long
    Dim result() As Variant
    
    localLowerBound = LBound(arr)
    localUpperBound = UBound(arr)
    ReDim result(localLowerBound To localUpperBound)
    For i = localLowerBound To localUpperBound
        result(i) = getScalarRepresentation(arr(i))
    Next
    ensureScalar1DArray = result
End Function

' TODO: Use or remove
Private Function OneDimArrayToJagged(ByRef arr() As Variant) As Variant()
    Dim i As Long
    Dim localLowerBound As Long
    Dim localUpperBound As Long
    Dim result() As Variant
    
    localLowerBound = LBound(arr)
    localUpperBound = UBound(arr)
    ReDim result(localLowerBound To localUpperBound)
    For i = localLowerBound To localUpperBound
        result(i) = Array(arr(i))
    Next
    OneDimArrayToJagged = result
End Function

' TODO: Use or remove
Private Function Transpose_1DArray(ByRef arr() As Variant) As Variant()
    'TODO: need check to ensure 1d?
    Dim i As Long
    Dim j As Long
    Dim localUpperBound As Long
    Dim localLowerBound As Long
    Dim result() As Variant
    localUpperBound = UBound(arr)
    localLowerBound = LBound(arr)
    ReDim result(localLowerBound To (localUpperBound - localLowerBound + 1), _
                 localLowerBound To localLowerBound)
    j = localLowerBound
    For i = localLowerBound To localUpperBound
        result(j, localLowerBound) = arr(i)
        j = j + 1
    Next
    Transpose_1DArray = result
End Function

' Used by ToExcelRange()
Private Function Transpose_2DArray(ByRef sourceArray() As Variant) As Variant()
    Dim currentRow As Long
    Dim LowerBoundRow As Long
    Dim UpperBoundRow As Long
    Dim currentColumn As Long
    Dim LowerBoundCol As Long
    Dim UpperBoundCol As Long
    Dim result() As Variant
    
    LowerBoundCol = LBound(sourceArray, 1)
    UpperBoundCol = UBound(sourceArray, 1)
    LowerBoundRow = LBound(sourceArray, 2)
    UpperBoundRow = UBound(sourceArray, 2)
    ReDim result(LowerBoundRow To UpperBoundRow, LowerBoundCol To UpperBoundCol)
    For currentRow = LowerBoundRow To UpperBoundRow
        For currentColumn = LowerBoundCol To UpperBoundRow
            result(currentRow, currentColumn) = sourceArray(currentColumn, currentRow)
        Next
    Next
    Transpose_2DArray = result
End Function


' TODO: Use or remove
Private Function Transpose_ArrayOfArrays(ByRef sourceArray() As Variant) As Variant()
    Dim currentRow As Long
    Dim LowerBoundRow As Long
    Dim UpperBoundRow As Long
    Dim currentColumn As Long
    Dim LowerBoundCol As Long
    Dim UpperBoundCol As Long
    Dim result() As Variant
    
    LowerBoundCol = LBound(sourceArray)
    UpperBoundCol = UBound(sourceArray)
    LowerBoundRow = LBound(sourceArray(LowerBoundCol))
    UpperBoundRow = UBound(sourceArray(LowerBoundCol))
    ReDim result(LowerBoundRow To UpperBoundRow, LowerBoundCol To UpperBoundCol)
    For currentRow = LowerBoundRow To UpperBoundRow
        For currentColumn = LowerBoundCol To UpperBoundCol
            result(currentRow, currentColumn) = sourceArray(currentColumn)(currentRow)
        Next
    Next
    Transpose_ArrayOfArrays = result
End Function

' TODO: Use or remove
Private Function ExtractColumnFromMultiDimensionalArray(ByRef sourceArray() As Variant, Optional ByVal col As Long = 1) As Variant()
    Dim i As Long
    Dim localLowerBound As Long
    Dim localUpperBound As Long
    Dim result() As Variant
    
    localLowerBound = LBound(sourceArray)
    localUpperBound = UBound(sourceArray)
    ReDim result(localLowerBound To localUpperBound)
    For i = localLowerBound To localUpperBound
        result(i) = sourceArray(i, col)
    Next
    ExtractColumnFromMultiDimensionalArray = result
End Function

Private Function getEmptyArray() As Variant()
    Dim result() As Variant
    ReDim result(pLowerBound)
    getEmptyArray = result
End Function

Private Sub quickSort(ByRef arr() As Variant, _
                      ByVal low As Long, _
                      ByVal high As Long, _
                      ByVal ArrayType As ArrayTypes, _
                      Optional ByVal col As Long)
    Dim partitionIndex As Long
    If low < high Then
        partitionIndex = qsPartition(arr, low, high, ArrayType, col)
        quickSort arr, low, partitionIndex - 1, ArrayType, col
        quickSort arr, partitionIndex + 1, high, ArrayType, col
    End If
End Sub

Private Function qsPartition(ByRef arr() As Variant, _
                             ByVal low As Long, _
                             ByVal high As Long, _
                             ByVal ArrayType As ArrayTypes, _
                             Optional ByVal col As Long) As Long
    Dim i As Long
    Dim j As Long
    Dim pivot As Variant
    Dim current As Variant
    
    If ArrayType = Jagged Then
        pivot = arr(high)(col)
    Else
        pivot = arr(high)
    End If
    
    i = low - 1
    For j = low To high - 1
        If ArrayType = Jagged Then
            current = arr(j)(col)
        Else
            current = arr(j)
        End If
    
        If current <= pivot Then
            i = i + 1
            Swap arr, i, j
        End If
    Next
    Swap arr, i + 1, high
    qsPartition = i + 1
End Function

Private Sub Swap(ByRef arr() As Variant, ByVal i As Long, ByVal j As Long)
    Dim temporaryArray As Variant
    temporaryArray = arr(i)
    arr(i) = arr(j)
    arr(j) = temporaryArray
End Sub

Private Function getInternalItems() As Variant()
    Dim result() As Variant
    result = pItems
    If Me.Capacity > pLength Then
        If pLength > 0 Then ReDim Preserve result(pLowerBound To Me.UpperBound)
    End If
    getInternalItems = result
End Function

Private Function recursiveToString(ByRef arr() As Variant, _
                                   ByVal PrettyPrint As Boolean, _
                                   Optional ByVal tabs As Long = 1) As String
    Const tabWidth As Long = 2
    Dim result As String
    Dim sep As String
    Dim i As Long

    result = "["
    For i = LBound(arr) To UBound(arr)
        sep = IIf(i = UBound(arr), "]", ", ")
        If IsArray(arr(i)) Then
            Dim nested() As Variant
            nested = arr(i)
            If PrettyPrint Then result = result & vbCrLf & Space(tabWidth * tabs)
            result = result & recursiveToString(nested, PrettyPrint, tabs + 1)
            result = result & IIf(i = UBound(arr) And PrettyPrint, vbCrLf & Space(tabWidth * (tabs - 1)) & sep, sep)
        Else
            result = result & Replace( _
                    CStr(getScalarRepresentation(arr(i))), _
                    ",", _
                    vbNullString _
                ) & sep
        End If
    Next
    recursiveToString = result
End Function

Private Function getScalarRepresentation(ByRef element As Variant) As Variant
    Dim result As Variant
    If IsObject(element) Then
        On Error Resume Next
        'Try to retrieve default member
        result = element
        On Error GoTo 0
        ' if there is no default member or it returns an object result = Empty
        If IsEmpty(result) Then
            result = "OBJECT"
        End If
    Else
        result = element
    End If
    ' if the element is array or default member returns an array then convert tostring
    If IsArray(result) And Not IsEmpty(result) Then
        Dim passThruArray() As Variant
        passThruArray = element
        result = recursiveToString(passThruArray, False)
    End If
    getScalarRepresentation = result
End Function

Private Function getArrayType(ByVal arr As Variant) As ArrayTypes
    Dim result As ArrayTypes
    If Not isArrayAllocated(arr) Then
        result = ArrayTypes.Unallocated
    Else
        If isMultidimensionalArray(arr) Then
            result = ArrayTypes.MultiDimension
        ElseIf isJaggedArray(arr) Then
            result = ArrayTypes.Jagged
        Else
            result = ArrayTypes.OneDimension
        End If
    End If
    getArrayType = result
End Function

Private Function isArrayAllocated(ByVal arr As Variant) As Boolean
    On Error Resume Next
    isArrayAllocated = IsArray(arr) And _
                                    Not IsError(LBound(arr, 1)) And _
                                    LBound(arr, 1) <= UBound(arr, 1)
    On Error GoTo 0
End Function

Private Function isJaggedArray(ByVal arr As Variant) As Boolean
    If IsArray(arr) Then
        On Error GoTo errHandler
        Dim element As Variant
        For Each element In arr
            If IsArray(element) Then
                isJaggedArray = True
                Exit Function
            End If
        Next
        On Error GoTo 0
    End If
    Exit Function
errHandler:
    Err.Clear
End Function

Private Function isMultidimensionalArray(ByVal arr As Variant) As Boolean
    If IsArray(arr) Then
        On Error GoTo errHandler
        '@Ignore VariableNotUsed
        Dim localUpperBound As Long
        '@Ignore AssignmentNotUsed
        localUpperBound = UBound(arr, 2)
        isMultidimensionalArray = True
        On Error GoTo 0
    End If
    Exit Function
errHandler:
    Err.Clear
End Function

'@Ignore FunctionReturnValueNotUsed
Private Function Rebase(Optional ByRef recurse As Boolean, Optional ByVal arr As Variant) As Variant()
    Dim isMD As Boolean
    Dim applyToMe As Boolean
    Dim newItems() As Variant
    
    If IsMissing(arr) Or Not IsArray(arr) Then applyToMe = True
    newItems = IIf(applyToMe, pItems, arr)
    isMD = isMultidimensionalArray(newItems)
    If isMD Then
        newItems = multiToJagged(newItems)
        recurse = True
    End If
    newItems = recursiveRebase(newItems, recurse)
    If isMD Then newItems = jaggedToMulti(newItems)
    If applyToMe Then pItems = newItems
    Rebase = newItems
End Function

Private Function recursiveRebase(ByRef arr() As Variant, ByVal recurse As Boolean) As Variant()
    Dim i As Long
    Dim localLowerBound As Long
    Dim localUpperBound As Long
    Dim offset As Long
    Dim newItems() As Variant
    
    localLowerBound = LBound(arr)
    localUpperBound = UBound(arr)
    offset = pLowerBound - localLowerBound
    ReDim newItems(pLowerBound To localUpperBound + offset)
    For i = localLowerBound To localUpperBound
        If IsArray(arr(i)) And recurse Then
            Dim nested() As Variant
            nested = arr(i)
            newItems(i + offset) = recursiveRebase(nested, recurse)
        Else
            newItems(i + offset) = arr(i)
        End If
    Next
    
    recursiveRebase = newItems
End Function

Private Sub EnsureCapacity(ByVal MinimumCapacity As Long)
    If Me.Capacity < MinimumCapacity Then
        Dim newCapacity As Long
        newCapacity = IIf(Me.Capacity = 0, DEFAULT_CAPACITY, Me.Capacity * 2)
            If newCapacity > MAX_ARRAY_LENGTH Then newCapacity = MAX_ARRAY_LENGTH
        If newCapacity < MinimumCapacity Then newCapacity = MinimumCapacity
        Me.Capacity = newCapacity
    End If
End Sub

Private Function copyRange(ByRef newItems() As Variant) As Variant()
    Dim i As Long
    For i = LBound(newItems) To UBound(newItems)
        newItems(i) = pItems(i)
    Next
    copyRange = newItems
End Function


Private Function jaggedToMulti(ByRef arr() As Variant, _
                      Optional ByVal depth As Long, _
                      Optional ByVal currentDepth As Long, _
                      Optional ByRef crumbs As Variant, _
                      Optional ByRef result As Variant, _
                      Optional ByVal EnsureScalar As Boolean) As Variant()
    Dim i As Long
    Dim localLowerBound As Long
    Dim localUpperBound As Long
    Dim localDepth As Long
    Dim localCurrentDepth As Long
    Dim localCrumbs() As Variant
    Dim localResult() As Variant
    
    
    localDepth = depth
    localCurrentDepth = currentDepth
    
    If localDepth = 0 Then
        If Not isJaggedArray(arr) Then
            Debug.Print "Argument passed to 'jaggedToMulti' not a jagged array"
            Exit Function
        End If
        localDepth = getJaggedArrayDepth(arr)
    End If
    
    If localDepth > 20 Then
        ' too many dimensions to be handled by getElementByBreadcrumb
        RaiseError MaxDimensionsLimit, "jaggedToMulti()", "arr()"
    ElseIf localDepth > 1 Then
    
        If IsMissing(crumbs) Then
            ReDim localCrumbs(localDepth - 1)
        Else
            localCrumbs = crumbs
        End If
        
        If IsMissing(result) Then
            localResult = createMultidimensionalArray(mapJaggedArray(arr, knownDepth:=localDepth))
        Else
            localResult = result
        End If
        
        localCurrentDepth = localCurrentDepth + 1
        localLowerBound = LBound(arr)
        localUpperBound = UBound(arr)
        For i = localLowerBound To localUpperBound
            localCrumbs(localCurrentDepth - 1) = i
            If localCurrentDepth = localDepth Or (Not IsArray(arr(i)) And Not EnsureScalar) Then
                If EnsureScalar Then
                    localResult = letElementByBreadcrumb(localResult, localCrumbs, getScalarRepresentation(arr(i)))
                Else
                    localResult = letElementByBreadcrumb(localResult, localCrumbs, arr(i))
                End If
            Else
                Dim nested() As Variant
                nested = arr(i)
                localResult = jaggedToMulti( _
                    arr:=nested, _
                    depth:=localDepth, _
                    currentDepth:=localCurrentDepth, _
                    crumbs:=localCrumbs, _
                    result:=localResult, _
                    EnsureScalar:=EnsureScalar _
                )
            End If
        Next
    Else
        localResult = arr
    End If
    
    jaggedToMulti = localResult
End Function

Private Function multiToJagged(ByRef arr() As Variant, _
                      Optional ByVal depth As Long, _
                      Optional ByVal currentDepth As Long, _
                      Optional ByRef crumbs As Variant) As Variant()
    Dim i As Long
    Dim localLowerBound As Long
    Dim localUpperBound As Long
    Dim localDepth As Long
    Dim localCurrentDepth As Long
    Dim localCrumbs() As Variant
    Dim result() As Variant
    
    localDepth = depth
    localCurrentDepth = currentDepth
    
    If localDepth = 0 Then
        If Not isMultidimensionalArray(arr) Then
            Debug.Print "Argument passed to 'multiToJagged' not a multidimensional array"
            Exit Function
        End If
        localDepth = getMultidimensionalArrayDepth(arr)
    End If
    
    If localDepth > 20 Then
        ' too many dimensions to be handled by getElementByBreadcrumb
        ' handle error
        Debug.Print "Cannot convert arrays with more than 20 dimensions to a jagged array!"
    ElseIf localDepth > 1 Then
        If IsMissing(crumbs) Then
            ReDim localCrumbs(localDepth - 1)
        Else
            localCrumbs = crumbs
        End If
        localCurrentDepth = localCurrentDepth + 1
        localLowerBound = LBound(arr, localCurrentDepth)
        localUpperBound = UBound(arr, localCurrentDepth)
        ReDim result(localLowerBound To localUpperBound)
        For i = localLowerBound To localUpperBound
            localCrumbs(localCurrentDepth - 1) = i
            If localCurrentDepth = localDepth Then
                result(i) = getElementByBreadcrumb(arr, localCrumbs)
            Else
                result(i) = multiToJagged(arr, localDepth, localCurrentDepth, localCrumbs)
            End If
        Next
    Else
        result = arr
    End If
    multiToJagged = result
End Function

Private Function getJaggedArrayDepth(ByRef arr() As Variant) As Long
    Dim i As Long
    Dim localLowerBound As Long
    Dim localUpperBound As Long
    Dim currentDepth As Long
    Dim maxDepth As Long
    Dim depth As Long
    
    If IsArray(arr) Then
        depth = depth + 1
        localLowerBound = LBound(arr)
        localUpperBound = UBound(arr)
        For i = localLowerBound To localUpperBound
            If IsArray(arr(i)) Then
                Dim nested() As Variant
                nested = arr(i)
                currentDepth = getJaggedArrayDepth(nested)
            End If
            If currentDepth > maxDepth Then maxDepth = currentDepth
        Next
        depth = depth + maxDepth
    End If
    getJaggedArrayDepth = depth
End Function

Private Function mapJaggedArray(ByRef arr() As Variant, _
                       Optional ByRef map As Variant, _
                       Optional ByVal currentIndex As Long, _
                       Optional ByVal knownDepth As Long) As Variant()
    Dim i As Long
    Dim localMap() As Variant
    Dim bounds() As Variant
    Dim element As Variant
    Dim localDepth As Long
    
    If Not IsArray(arr) Then Exit Function
    If IsMissing(map) Then
        If knownDepth > 0 Then
            localDepth = knownDepth
        Else
            localDepth = getJaggedArrayDepth(arr)
        End If
        ReDim localMap(0 To localDepth - 1)
        For i = LBound(localMap) To UBound(localMap)
            localMap(i) = Array(0, 0)
        Next
    Else
        localMap = map
    End If
    
    bounds = getArrayBounds(arr)
    
    If bounds(0) < localMap(currentIndex)(0) Then localMap(currentIndex)(0) = bounds(0)
    If bounds(1) > localMap(currentIndex)(1) Then localMap(currentIndex)(1) = bounds(1)
    
    If currentIndex < localDepth Then
        For Each element In arr
            If IsArray(element) Then
                Dim nested() As Variant
                nested = element
                localMap = mapJaggedArray(nested, localMap, currentIndex + 1)
            End If
        Next
    End If
    
    mapJaggedArray = localMap
End Function

Private Function getArrayBounds(ByVal arr As Variant) As Variant()
    Dim result(0 To 1) As Variant
    If IsArray(arr) Then
        result(0) = LBound(arr)
        result(1) = UBound(arr)
    End If
    getArrayBounds = result
End Function

Private Function getMultidimensionalArrayDepth(ByVal arr As Variant) As Long
    Dim i As Long
    '@Ignore VariableNotUsed
    Dim void As Long
    
    On Error Resume Next
    Do
        i = i + 1
        void = UBound(arr, i)
    Loop Until Err.number <> 0
    Err.Clear
    On Error GoTo 0
    getMultidimensionalArrayDepth = i - 1
End Function

Private Function getElementByBreadcrumb(ByRef arr() As Variant, ByRef crumb() As Variant) As Variant
    Dim result As Variant
    Select Case UBound(crumb)
    Case 0
        result = arr(crumb(0))
    Case 1
        result = arr(crumb(0), crumb(1))
    Case 2
        result = arr(crumb(0), crumb(1), crumb(2))
    Case 3
        result = arr(crumb(0), crumb(1), crumb(2), crumb(3))
    Case 4
        result = arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4))
    Case 5
        result = arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5))
    Case 6
        result = arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6))
    Case 7
        result = arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7))
    Case 8
        result = arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8))
    Case 9
        result = arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9))
    Case 10
        result = arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10))
    Case 11
        result = arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                     crumb(11))
    Case 12
        result = arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                     crumb(11), crumb(12))
    Case 13
        result = arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                     crumb(11), crumb(12), crumb(13))
    Case 14
        result = arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                     crumb(11), crumb(12), crumb(13), crumb(14))
    Case 15
        result = arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                     crumb(11), crumb(12), crumb(13), crumb(14), crumb(15))
    Case 16
        result = arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                     crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16))
    Case 17
        result = arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                     crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17))
    Case 18
        result = arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                     crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18))
    Case 19
        result = arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                     crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18), crumb(19))
    Case 20
        result = arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                     crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18), crumb(19), crumb(20))
    End Select
    getElementByBreadcrumb = result
End Function

Private Function letElementByBreadcrumb(ByRef arr() As Variant, ByRef crumb() As Variant, ByVal element As Variant) As Variant
    If IsObject(element) Then
        Select Case UBound(crumb)
        Case 0
            Set arr(crumb(0)) = element
        Case 1
            Set arr(crumb(0), crumb(1)) = element
        Case 2
            Set arr(crumb(0), crumb(1), crumb(2)) = element
        Case 3
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3)) = element
        Case 4
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4)) = element
        Case 5
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5)) = element
        Case 6
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6)) = element
        Case 7
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7)) = element
        Case 8
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8)) = element
        Case 9
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9)) = element
        Case 10
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10)) = element
        Case 11
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                    crumb(11)) = element
        Case 12
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                    crumb(11), crumb(12)) = element
        Case 13
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                    crumb(11), crumb(12), crumb(13)) = element
        Case 14
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                    crumb(11), crumb(12), crumb(13), crumb(14)) = element
        Case 15
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                    crumb(11), crumb(12), crumb(13), crumb(14), crumb(15)) = element
        Case 16
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                    crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16)) = element
        Case 17
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                    crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17)) = element
        Case 18
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                    crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18)) = element
        Case 19
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                    crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18), crumb(19)) = element
        Case 20
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                    crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18), crumb(19), crumb(20)) = element
        End Select
    Else
        Select Case UBound(crumb)
        Case 0
            arr(crumb(0)) = element
        Case 1
            arr(crumb(0), crumb(1)) = element
        Case 2
            arr(crumb(0), crumb(1), crumb(2)) = element
        Case 3
            arr(crumb(0), crumb(1), crumb(2), crumb(3)) = element
        Case 4
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4)) = element
        Case 5
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5)) = element
        Case 6
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6)) = element
        Case 7
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7)) = element
        Case 8
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8)) = element
        Case 9
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9)) = element
        Case 10
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10)) = element
        Case 11
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                crumb(11)) = element
        Case 12
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                crumb(11), crumb(12)) = element
        Case 13
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                crumb(11), crumb(12), crumb(13)) = element
        Case 14
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                crumb(11), crumb(12), crumb(13), crumb(14)) = element
        Case 15
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                crumb(11), crumb(12), crumb(13), crumb(14), crumb(15)) = element
        Case 16
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16)) = element
        Case 17
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17)) = element
        Case 18
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18)) = element
        Case 19
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18), crumb(19)) = element
        Case 20
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18), crumb(19), crumb(20)) = element
        End Select
    End If
    letElementByBreadcrumb = arr
End Function

Private Function createMultidimensionalArray(ByRef crumb() As Variant) As Variant()
    Dim result() As Variant
    Select Case UBound(crumb)
    Case 0
        ReDim result(crumb(0)(0) To crumb(0)(1))
    Case 1
        ReDim result(crumb(0)(0) To crumb(0)(1), crumb(1)(0) To crumb(1)(1))
    Case 2
        ReDim result(crumb(0)(0) To crumb(0)(1), crumb(1)(0) To crumb(1)(1), crumb(2)(0) To crumb(2)(1))
    Case 3
        ReDim result(crumb(0)(0) To crumb(0)(1), crumb(1)(0) To crumb(1)(1), crumb(2)(0) To crumb(2)(1), crumb(3)(0) To crumb(3)(1))
    Case 4
        ReDim result(crumb(0)(0) To crumb(0)(1), crumb(1)(0) To crumb(1)(1), crumb(2)(0) To crumb(2)(1), crumb(3)(0) To crumb(3)(1), _
                     crumb(4)(0) To crumb(4)(1))
    Case 5
        ReDim result(crumb(0)(0) To crumb(0)(1), crumb(1)(0) To crumb(1)(1), crumb(2)(0) To crumb(2)(1), crumb(3)(0) To crumb(3)(1), _
                     crumb(4)(0) To crumb(4)(1), crumb(5)(0) To crumb(5)(1))
    Case 6
        ReDim result(crumb(0)(0) To crumb(0)(1), crumb(1)(0) To crumb(1)(1), crumb(2)(0) To crumb(2)(1), crumb(3)(0) To crumb(3)(1), _
                     crumb(4)(0) To crumb(4)(1), crumb(5)(0) To crumb(5)(1), crumb(6)(0) To crumb(6)(1))
    Case 7
        ReDim result(crumb(0)(0) To crumb(0)(1), crumb(1)(0) To crumb(1)(1), crumb(2)(0) To crumb(2)(1), crumb(3)(0) To crumb(3)(1), _
                     crumb(4)(0) To crumb(4)(1), crumb(5)(0) To crumb(5)(1), crumb(6)(0) To crumb(6)(1), crumb(7)(0) To crumb(7)(1))
    Case 8
        ReDim result(crumb(0)(0) To crumb(0)(1), crumb(1)(0) To crumb(1)(1), crumb(2)(0) To crumb(2)(1), crumb(3)(0) To crumb(3)(1), _
                     crumb(4)(0) To crumb(4)(1), crumb(5)(0) To crumb(5)(1), crumb(6)(0) To crumb(6)(1), crumb(7)(0) To crumb(7)(1), _
                     crumb(8)(0) To crumb(8)(1))
    Case 9
        ReDim result(crumb(0)(0) To crumb(0)(1), crumb(1)(0) To crumb(1)(1), crumb(2)(0) To crumb(2)(1), crumb(3)(0) To crumb(3)(1), _
                     crumb(4)(0) To crumb(4)(1), crumb(5)(0) To crumb(5)(1), crumb(6)(0) To crumb(6)(1), crumb(7)(0) To crumb(7)(1), _
                     crumb(8)(0) To crumb(8)(1), crumb(9)(0) To crumb(9)(1))
    Case 10
        ReDim result(crumb(0)(0) To crumb(0)(1), crumb(1)(0) To crumb(1)(1), crumb(2)(0) To crumb(2)(1), crumb(3)(0) To crumb(3)(1), _
                     crumb(4)(0) To crumb(4)(1), crumb(5)(0) To crumb(5)(1), crumb(6)(0) To crumb(6)(1), crumb(7)(0) To crumb(7)(1), _
                     crumb(8)(0) To crumb(8)(1), crumb(9)(0) To crumb(9)(1), crumb(10)(0) To crumb(10)(1))
    Case 11
        ReDim result(crumb(0)(0) To crumb(0)(1), crumb(1)(0) To crumb(1)(1), crumb(2)(0) To crumb(2)(1), crumb(3)(0) To crumb(3)(1), _
                     crumb(4)(0) To crumb(4)(1), crumb(5)(0) To crumb(5)(1), crumb(6)(0) To crumb(6)(1), crumb(7)(0) To crumb(7)(1), _
                     crumb(8)(0) To crumb(8)(1), crumb(9)(0) To crumb(9)(1), crumb(10)(0) To crumb(10)(1), crumb(11)(0) To crumb(11)(1))
    Case 12
        ReDim result(crumb(0)(0) To crumb(0)(1), crumb(1)(0) To crumb(1)(1), crumb(2)(0) To crumb(2)(1), crumb(3)(0) To crumb(3)(1), _
                     crumb(4)(0) To crumb(4)(1), crumb(5)(0) To crumb(5)(1), crumb(6)(0) To crumb(6)(1), crumb(7)(0) To crumb(7)(1), _
                     crumb(8)(0) To crumb(8)(1), crumb(9)(0) To crumb(9)(1), crumb(10)(0) To crumb(10)(1), crumb(11)(0) To crumb(11)(1), _
                     crumb(12)(0) To crumb(12)(1))
    Case 13
        ReDim result(crumb(0)(0) To crumb(0)(1), crumb(1)(0) To crumb(1)(1), crumb(2)(0) To crumb(2)(1), crumb(3)(0) To crumb(3)(1), _
                     crumb(4)(0) To crumb(4)(1), crumb(5)(0) To crumb(5)(1), crumb(6)(0) To crumb(6)(1), crumb(7)(0) To crumb(7)(1), _
                     crumb(8)(0) To crumb(8)(1), crumb(9)(0) To crumb(9)(1), crumb(10)(0) To crumb(10)(1), crumb(11)(0) To crumb(11)(1), _
                     crumb(12)(0) To crumb(12)(1), crumb(13)(0) To crumb(13)(1))
    Case 14
        ReDim result(crumb(0)(0) To crumb(0)(1), crumb(1)(0) To crumb(1)(1), crumb(2)(0) To crumb(2)(1), crumb(3)(0) To crumb(3)(1), _
                     crumb(4)(0) To crumb(4)(1), crumb(5)(0) To crumb(5)(1), crumb(6)(0) To crumb(6)(1), crumb(7)(0) To crumb(7)(1), _
                     crumb(8)(0) To crumb(8)(1), crumb(9)(0) To crumb(9)(1), crumb(10)(0) To crumb(10)(1), crumb(11)(0) To crumb(11)(1), _
                     crumb(12)(0) To crumb(12)(1), crumb(13)(0) To crumb(13)(1), crumb(14)(0) To crumb(14)(1))
    Case 15
        ReDim result(crumb(0)(0) To crumb(0)(1), crumb(1)(0) To crumb(1)(1), crumb(2)(0) To crumb(2)(1), crumb(3)(0) To crumb(3)(1), _
                     crumb(4)(0) To crumb(4)(1), crumb(5)(0) To crumb(5)(1), crumb(6)(0) To crumb(6)(1), crumb(7)(0) To crumb(7)(1), _
                     crumb(8)(0) To crumb(8)(1), crumb(9)(0) To crumb(9)(1), crumb(10)(0) To crumb(10)(1), crumb(11)(0) To crumb(11)(1), _
                     crumb(12)(0) To crumb(12)(1), crumb(13)(0) To crumb(13)(1), crumb(14)(0) To crumb(14)(1), crumb(15)(0) To crumb(15)(1))
    Case 16
        ReDim result(crumb(0)(0) To crumb(0)(1), crumb(1)(0) To crumb(1)(1), crumb(2)(0) To crumb(2)(1), crumb(3)(0) To crumb(3)(1), _
                     crumb(4)(0) To crumb(4)(1), crumb(5)(0) To crumb(5)(1), crumb(6)(0) To crumb(6)(1), crumb(7)(0) To crumb(7)(1), _
                     crumb(8)(0) To crumb(8)(1), crumb(9)(0) To crumb(9)(1), crumb(10)(0) To crumb(10)(1), crumb(11)(0) To crumb(11)(1), _
                     crumb(12)(0) To crumb(12)(1), crumb(13)(0) To crumb(13)(1), crumb(14)(0) To crumb(14)(1), crumb(15)(0) To crumb(15)(1), _
                     crumb(16)(0) To crumb(16)(1))
    Case 17
        ReDim result(crumb(0)(0) To crumb(0)(1), crumb(1)(0) To crumb(1)(1), crumb(2)(0) To crumb(2)(1), crumb(3)(0) To crumb(3)(1), _
                     crumb(4)(0) To crumb(4)(1), crumb(5)(0) To crumb(5)(1), crumb(6)(0) To crumb(6)(1), crumb(7)(0) To crumb(7)(1), _
                     crumb(8)(0) To crumb(8)(1), crumb(9)(0) To crumb(9)(1), crumb(10)(0) To crumb(10)(1), crumb(11)(0) To crumb(11)(1), _
                     crumb(12)(0) To crumb(12)(1), crumb(13)(0) To crumb(13)(1), crumb(14)(0) To crumb(14)(1), crumb(15)(0) To crumb(15)(1), _
                     crumb(16)(0) To crumb(16)(1), crumb(17)(0) To crumb(17)(1))
    Case 18
        ReDim result(crumb(0)(0) To crumb(0)(1), crumb(1)(0) To crumb(1)(1), crumb(2)(0) To crumb(2)(1), crumb(3)(0) To crumb(3)(1), _
                     crumb(4)(0) To crumb(4)(1), crumb(5)(0) To crumb(5)(1), crumb(6)(0) To crumb(6)(1), crumb(7)(0) To crumb(7)(1), _
                     crumb(8)(0) To crumb(8)(1), crumb(9)(0) To crumb(9)(1), crumb(10)(0) To crumb(10)(1), crumb(11)(0) To crumb(11)(1), _
                     crumb(12)(0) To crumb(12)(1), crumb(13)(0) To crumb(13)(1), crumb(14)(0) To crumb(14)(1), crumb(15)(0) To crumb(15)(1), _
                     crumb(16)(0) To crumb(16)(1), crumb(17)(0) To crumb(17)(1), crumb(18)(0) To crumb(18)(1))
    Case 19
        ReDim result(crumb(0)(0) To crumb(0)(1), crumb(1)(0) To crumb(1)(1), crumb(2)(0) To crumb(2)(1), crumb(3)(0) To crumb(3)(1), _
                     crumb(4)(0) To crumb(4)(1), crumb(5)(0) To crumb(5)(1), crumb(6)(0) To crumb(6)(1), crumb(7)(0) To crumb(7)(1), _
                     crumb(8)(0) To crumb(8)(1), crumb(9)(0) To crumb(9)(1), crumb(10)(0) To crumb(10)(1), crumb(11)(0) To crumb(11)(1), _
                     crumb(12)(0) To crumb(12)(1), crumb(13)(0) To crumb(13)(1), crumb(14)(0) To crumb(14)(1), crumb(15)(0) To crumb(15)(1), _
                     crumb(16)(0) To crumb(16)(1), crumb(17)(0) To crumb(17)(1), crumb(18)(0) To crumb(18)(1), crumb(19)(0) To crumb(19)(1))
    Case 20
        ReDim result(crumb(0)(0) To crumb(0)(1), crumb(1)(0) To crumb(1)(1), crumb(2)(0) To crumb(2)(1), crumb(3)(0) To crumb(3)(1), _
                     crumb(4)(0) To crumb(4)(1), crumb(5)(0) To crumb(5)(1), crumb(6)(0) To crumb(6)(1), crumb(7)(0) To crumb(7)(1), _
                     crumb(8)(0) To crumb(8)(1), crumb(9)(0) To crumb(9)(1), crumb(10)(0) To crumb(10)(1), crumb(11)(0) To crumb(11)(1), _
                     crumb(12)(0) To crumb(12)(1), crumb(13)(0) To crumb(13)(1), crumb(14)(0) To crumb(14)(1), crumb(15)(0) To crumb(15)(1), _
                     crumb(16)(0) To crumb(16)(1), crumb(17)(0) To crumb(17)(1), crumb(18)(0) To crumb(18)(1), crumb(19)(0) To crumb(19)(1), _
                     crumb(20)(0) To crumb(20)(1))
    End Select
    createMultidimensionalArray = result
End Function










