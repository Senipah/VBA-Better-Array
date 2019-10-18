VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BetterArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "An array class for VBA providing features found in more modern languages "
Option Explicit

'@Folder("BetterArray")
'@ModuleDescription("An array class for VBA providing features found in more modern languages ")

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
    Source As String
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

'@Description("Constructor")
Private Sub Class_Initialize()
Attribute Class_Initialize.VB_Description = "Constructor"
    Me.Capacity = DEFAULT_CAPACITY
    PopulateErrorDefinitions
End Sub


'@Ignore EmptyMethod
'@Description("Destructor")
Private Sub Class_Terminate()
Attribute Class_Terminate.VB_Description = "Destructor"

End Sub

'''''''''''''''''''''
' Public Properties '
'''''''''''''''''''''

'@Description("Returns the capacity of the internal array.")
Public Property Get Capacity() As Long
Attribute Capacity.VB_Description = "Returns the capacity of the internal array."
    If IsEmpty(pCapacity) Then Me.Capacity = DEFAULT_CAPACITY
    Capacity = pCapacity
End Property

'@Description("Sets the capacity of the internal array.")
Public Property Let Capacity(ByVal Value As Long)
Attribute Capacity.VB_Description = "Sets the capacity of the internal array."
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

'@Description("Returns the number of entries stored in the array.")
Public Property Get Length() As Long
Attribute Length.VB_Description = "Returns the number of entries stored in the array."
    Length = pLength
End Property

'@Description("Returns largest used index in the array.")
Public Property Get UpperBound()
Attribute UpperBound.VB_Description = "Returns largest used index in the array."
    UpperBound = pLength + pLowerBound - 1
End Property

'@Description("Returns smallest used index in the array.")
Public Property Get LowerBound() As Long
Attribute LowerBound.VB_Description = "Returns smallest used index in the array."
    LowerBound = pLowerBound
End Property

'@Description("Sets the starting index of the array. If an array is already stored then it will be re-indexed to match the new value")
Public Property Let LowerBound(ByVal Value As Long)
Attribute LowerBound.VB_Description = "Sets the starting index of the array. If an array is already stored then it will be re-indexed to match the new value"
    If Value <> pLowerBound And Value >= 0 Then
        pLowerBound = Value
        Rebase
    End If
End Property

'@DefaultMember
'@Description("Returns the element in the array stored at the specified index.")
Public Property Get Item(ByVal index As Long)
Attribute Item.VB_Description = "Returns the element in the array stored at the specified index."
Attribute Item.VB_UserMemId = 0
    If index <= pLength Then
        ' not using letOrSetElement here as it might try to access the Items letter
        If IsObject(pItems(index)) Then
            Set Item = pItems(index)
        Else
            Item = pItems(index)
        End If
    Else
        Err.Raise 9
    End If
End Property

'@Description("Assigns the passed element to the array at the specified index.")
Public Property Let Item(ByVal index As Long, ByVal Element As Variant)
Attribute Item.VB_Description = "Assigns the passed element to the array at the specified index."
    If Me.UpperBound >= index Then
        If index < Me.LowerBound Then
            Me.Unshift Element
        Else
            letOrSetElement pItems(index), Element
        End If
    Else
        Me.Push Element
    End If
End Property

'@Description("Returns the stored array.")
Public Property Get Items() As Variant
Attribute Items.VB_Description = "Returns the stored array."
    Dim result() As Variant
    result = getInternalItems()
    If pArrayType = ArrayTypes.MultiDimension And isJaggedArray(result) Then result = jaggedToMulti(result)
    Items = result
End Property

'@Description("Stores the passed array.")
Public Property Let Items(ByVal Values As Variant)
Attribute Items.VB_Description = "Stores the passed array."
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
'@Description("Adds the provided argments to the array at the next available index.")
Public Function Push(ParamArray args() As Variant) As Variant
Attribute Push.VB_Description = "Adds the provided argments to the array at the next available index."
    Dim Element As Variant

    If pArrayType = ArrayTypes.Unallocated Or _
       pArrayType = ArrayTypes.Undefined Then
        pArrayType = ArrayTypes.OneDimension
    End If

    For Each Element In args
        If IsArray(Element) And pArrayType = ArrayTypes.OneDimension Then
            pArrayType = ArrayTypes.Jagged
        End If
        If pLength = pCapacity Then
            EnsureCapacity pLength + 1
        End If
        letOrSetElement pItems(pLength), Element
        pLength = pLength + 1
    Next
    Push = pLength
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Removes the last element from the array and returns that element.")
Public Function Pop() As Variant
Attribute Pop.VB_Description = "Removes the last element from the array and returns that element."
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
'@Description("Removes the first element from the array and returns that element.")
Public Function Shift() As Variant
Attribute Shift.VB_Description = "Removes the first element from the array and returns that element."
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
'@Description("Adds one or more elements to the beginning of an array. Returns the new length of the array.")
Public Function Unshift(ParamArray args() As Variant) As Long
Attribute Unshift.VB_Description = "Adds one or more elements to the beginning of an array. Returns the new length of the array."
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

' TODO: Document
'@Ignore FunctionReturnValueNotUsed
'@Description("Parses a string representing an array and stores the outcome.")
Public Function ParseFromString(ByVal SourceString As String, _
                                Optional ByVal ValueSeparator As String = ",", _
                                Optional ByVal ArrayOpenDelimiter As String, _
                                Optional ByVal ArrayClosingDelimiter As String) As BetterArray
Attribute ParseFromString.VB_Description = "Parses a string representing an array and stores the outcome."
    Dim opener As String
    Dim closer As String
    Dim ArraysAreDelimited As Boolean
    Dim result() As Variant
    
    If Len(SourceString) > 0 Then
        If ArrayOpenDelimiter = vbNullString And ArrayClosingDelimiter = vbNullString Then
            Dim firstChar As String
            Dim lastChar As String
            firstChar = Left$(SourceString, 1)
            lastChar = Right$(SourceString, 1)
            ' no array delimiters supplied
            ' see if the diff in ascii values are for first and last chars are less than 3
            ' this covers () [] {} <>
            ' escape numbers incase array is "1,2,3"
            ' escape ucase letters incase array is "A,B,C"
            ' escape lcase letters incase array is "a,b,c"
            If (Asc(lastChar) - Asc(firstChar) < 3) And _
                    (Asc(firstChar) < 65 Or Asc(firstChar) > 90) And _
                    (Asc(firstChar) < 97 Or Asc(firstChar) > 122) And _
                    Not IsNumeric(firstChar) Then
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
            result = ParseDelimitedArrayString(SourceString, ValueSeparator, opener, closer)
        Else
            result = ParseArraySegmentFromString(SourceString, 1, 0, ValueSeparator)
        End If
    Else
        ' nullstring passed
        ' handle exception
    End If
    
    Me.Items = result
    Set ParseFromString = Me
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Returns a string representing the array structure and its elements.")
Public Function ToString(Optional ByVal PrettyPrint As Boolean, _
                         Optional ByVal DelimitStrings As Boolean) As String
Attribute ToString.VB_Description = "Returns a string representing the array structure and its elements."
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
        result = recursiveToString(localItems, PrettyPrint, DelimitStrings:=DelimitStrings)
    End If
    ToString = result
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Determines whether the array includes a certain value among its entries.")
Public Function Includes(ByVal SearchElement As Variant, Optional ByVal FromIndex As Long) As Boolean
Attribute Includes.VB_Description = "Determines whether the array includes a certain value among its entries."
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
'@Description("Returns a new array that contains the keys for each index in the array.")
Public Function Keys() As Variant()
Attribute Keys.VB_Description = "Returns a new array that contains the keys for each index in the array."
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
'@Description("Returns the largest value in a list of values. If arguments are provided it will return the largest argument. If no arguments are provided it will return the largest element in the stored array")
Public Function Max(ParamArray args() As Variant) As Variant
Attribute Max.VB_Description = "Returns the largest value in a list of values. If arguments are provided it will return the largest argument. If no arguments are provided it will return the largest element in the stored array"
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
'@Description("Returns the smallest value in a list of values. If arguments are provided it will return the smallest argument. If no arguments are provided it will return the smallest element in the stored array")
Public Function Min(ParamArray args() As Variant) As Variant
Attribute Min.VB_Description = "Returns the smallest value in a list of values. If arguments are provided it will return the smallest argument. If no arguments are provided it will return the smallest element in the stored array"
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
'@Description("Returns a new array containing the portion of the stored array between the StartIndex and EndIndex")
Public Function Slice(ByVal StartIndex As Long, Optional ByVal EndIndex As Long) As Variant()
Attribute Slice.VB_Description = "Returns a new array containing the portion of the stored array between the StartIndex and EndIndex"
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

' TODO: Implement, Document & Test FromExcelRange method
'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Stores the values contained in an Excel range in the internal array")
Public Function FromExcelRange(ByRef FromRange As Object, _
                               Optional ByVal ColumnNumber As Long, _
                               Optional ByVal RowNumber As Long) As BetterArray
Attribute FromExcelRange.VB_Description = "Stores the values contained in an Excel range in the internal array"
    Dim localColumnNumber As Long
    Dim localRowNumber As Long
    
    If TypeName(FromRange) = "Range" Then
        If FromRange.Columns.count = 1 Then
            localColumnNumber = 1
        Else
            localColumnNumber = ColumnNumber
        End If
        
        If FromRange.Rows.count = 1 Then
            localRowNumber = 1
        Else
            localRowNumber = RowNumber
        End If
        
        Me.Items = FromRange.Value
    
        If localColumnNumber > 0 Then
            'Extract 1d array based on column number
            Me.Items = Me.ExtractColumn(localColumnNumber)
        ElseIf localRowNumber > 0 Then
            'Extract 1d array based on row number
            Me.Items = Me.ExtractRow(localRowNumber)
        End If
    Else
        RaiseError ErrorCodes.ExpectedRangeObject, "FromExcelRange()", "FromRange"
    End If
    
    Set FromExcelRange = Me
End Function

'@Description("Extracts the specified column of a multidimensional/jagged array as a new array.")
Public Function ExtractColumn(Optional ByVal ColumnNumber As Long) As Variant()
Attribute ExtractColumn.VB_Description = "Extracts the specified column of a multidimensional/jagged array as a new array."
    Dim i As Long
    Dim localColumnNumber As Long
    Dim localLowerBound As Long
    Dim localUpperBound As Long
    Dim sourceArray() As Variant
    Dim result() As Variant
    
    sourceArray = getInternalItems()
    If pArrayType = ArrayTypes.MultiDimension Or pArrayType = ArrayTypes.Jagged Then
        localLowerBound = LBound(sourceArray)
        localUpperBound = UBound(sourceArray)
        ReDim result(localLowerBound To localUpperBound)
        
        If ColumnNumber >= localLowerBound And ColumnNumber <= localUpperBound Then
            localColumnNumber = ColumnNumber
        Else
            localColumnNumber = Me.LowerBound
        End If
        
        For i = localLowerBound To localUpperBound
            letOrSetElement result(i), sourceArray(i)(localColumnNumber)
        Next
    Else
        ' can't extract from 1d or empty - return unchanged list
        result = sourceArray
    End If
    
    ExtractColumn = result
End Function

'@Description("Extracts the specified row of a multidimensional/jagged array as a new array.")
Public Function ExtractRow(Optional ByVal RowNumber As Long) As Variant()
Attribute ExtractRow.VB_Description = "Extracts the specified row of a multidimensional/jagged array as a new array."
    Dim localRowNumber As Long
    Dim localLowerBound As Long
    Dim localUpperBound As Long
    Dim sourceArray() As Variant
    Dim result() As Variant

    sourceArray = getInternalItems()
    If pArrayType = ArrayTypes.MultiDimension Or pArrayType = ArrayTypes.Jagged Then
        localLowerBound = LBound(sourceArray)
        localUpperBound = UBound(sourceArray)
        ReDim result(localLowerBound To localUpperBound)

        If RowNumber >= localLowerBound And RowNumber <= localUpperBound Then
            localRowNumber = RowNumber
        Else
            localRowNumber = Me.LowerBound
        End If
        letOrSetElement result, sourceArray(localRowNumber)
    Else
        ' can't extract from 1d or empty - return
        result = sourceArray
    End If
    ExtractRow = result
End Function

' TODO: Implement, Document & Test ToExcelRange method
'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Writes the values stored in the array to an excel worksheet starting at the specified range.")
Public Function ToExcelRange(ByRef Destination As Object, Optional ByVal TransposeValues As Boolean) As Object
Attribute ToExcelRange.VB_Description = "Writes the values stored in the array to an excel worksheet starting at the specified range."
    Dim localRange As Object
    Dim localItems() As Variant
    Dim depth As Long
    Dim lengthRows As Long
    Dim lengthColumns As Long
    
    lengthRows = 1 'default
    lengthColumns = 1 'default
    
    If TypeName(Destination) = "Range" Then
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
            Set localRange = Destination.Resize( _
                RowSize:=lengthRows, _
                ColumnSize:=lengthColumns _
            )
            localRange.Value = localItems
        Else
            'array must be unalloc
            'do nothing(?)
        End If
    Else
        RaiseError ErrorCodes.ExpectedRangeObject, "ToExcelRange()", "Destination"
    End If
    
    Set ToExcelRange = localRange
End Function

''''''''''''''''''''''''''''
' Chainable Public Methods '
''''''''''''''''''''''''''''

' TODO: Document & Test - Clone
'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Returns a new BetterArray instance containing the same values as the current instance.")
Public Function Clone() As BetterArray
Attribute Clone.VB_Description = "Returns a new BetterArray instance containing the same values as the current instance."
    Dim result As BetterArray
    Set result = New BetterArray
    result.LowerBound = Me.LowerBound
    result.Items = Me.Items
    Set Clone = result
End Function

' TODO: Document & Test - ResetToDefault
'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Clears the current array and resets capacity to default value")
Public Function ResetToDefault() As BetterArray
Attribute ResetToDefault.VB_Description = "Clears the current array and resets capacity to default value"
    pLowerBound = 0
    pItems = getEmptyArray()
    pLength = 0
    Me.Capacity = DEFAULT_CAPACITY
    Set ResetToDefault = Me
End Function

' TODO: Document & Test - ResetToDefault
'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Clears all entries in the current array but retains the same internal capacity")
Public Function Clear() As BetterArray
Attribute Clear.VB_Description = "Clears all entries in the current array but retains the same internal capacity"
    pItems = getEmptyArray()
    pLength = 0
    Me.Capacity = pCapacity
    Set Clear = Me
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Joins one or more arrays onto the end of the current array.")
Public Function Concat(ParamArray args() As Variant) As BetterArray
Attribute Concat.VB_Description = "Joins one or more arrays onto the end of the current array."
    Dim arg As Variant
    Dim Element As Variant
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
            For Each Element In localArray
                Me.Push Element
            Next
        Else
            Me.Push arg
        End If
    Next
    Set Concat = Me
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Writes all of the elements in a Collection objet to the internal array.")
Public Function CopyFromCollection(ByVal SourceCollection As Collection) As BetterArray
Attribute CopyFromCollection.VB_Description = "Writes all of the elements in a Collection objet to the internal array."
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
'@Description("Sorts the stored array in ascending order.")
Public Function Sort(Optional ByVal SortColumn As Long) As BetterArray
Attribute Sort.VB_Description = "Sorts the stored array in ascending order."
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
                    ' TODO: Raise custom err
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
'@Description("Copies a section of the stored array to another location in the same array.")
Public Function CopyWithin(ByVal Target As Long, _
                           Optional ByVal StartIndex As Long, _
                           Optional ByVal EndIndex As Long) As BetterArray
Attribute CopyWithin.VB_Description = "Copies a section of the stored array to another location in the same array."
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

' TODO: add support for 2d arrays. Consider rename?
'@Ignore ShadowedDeclaration
'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Filters the stored array to include only values that meet the specified criteria.")
Public Function Filter(ByVal Match As Variant, Optional ByVal Exclude As Boolean) As BetterArray
Attribute Filter.VB_Description = "Filters the stored array to include only values that meet the specified criteria."
    Dim Element As Variant
    Dim oldArr() As Variant
    Dim newArr() As Variant
    Dim i As Long
    
    oldArr = getInternalItems()
    ReDim newArr(LBound(oldArr), UBound(oldArr))
    i = Me.LowerBound
    For Each Element In oldArr
        If (Exclude And Element = Match) Or (Not Exclude And Element <> Match) Then
            newArr(i) = Element
            i = i + 1
        End If
    Next
    ReDim Preserve newArr(Me.LowerBound, i)
    Me.Items = newArr
    Set Filter = Me
End Function

' TODO: Check implementation against spec
'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Reverses the order of elements in the array. The first array element becomes the last, and the last array element becomes the first.")
Public Function Reverse() As BetterArray
Attribute Reverse.VB_Description = "Reverses the order of elements in the array. The first array element becomes the last, and the last array element becomes the first."
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

' TODO: Refactor to remove recursion
'@Description("Recursively parses nested arrays from a string.")
Private Function ParseDelimitedArrayString(ByVal SourceString As String, _
                              ByVal ValueSeparator As String, _
                              ByVal opener As String, _
                              ByVal closer As String, _
                              Optional ByRef cursor As Long = 2) As Variant()
Attribute ParseDelimitedArrayString.VB_Description = "Recursively parses nested arrays from a string."
    Dim currentChar As String
    Dim localResult() As Variant
    Dim i As Long
    Dim breakLoop As Boolean
    Dim nextOpener As Long
    Dim nextCloser As Long
    
    currentChar = Mid$(SourceString, cursor, 1)
    nextOpener = InStr(cursor, SourceString, opener)
    nextCloser = InStr(cursor, SourceString, closer)
    
    If currentChar <> opener And (nextCloser < nextOpener Or nextOpener = 0) Then
        ' we're inside an array to be parsed
        ParseDelimitedArrayString = ParseArraySegmentFromString(SourceString, cursor, nextCloser, ValueSeparator)
        Exit Function
    End If
        
    ' traverse nested arrays
    i = pLowerBound
    Do
        If currentChar = opener Then
            If i > pLowerBound Then
                ReDim Preserve localResult(i)
            Else
                localResult = getEmptyArray
            End If
            cursor = cursor + 1
            localResult(i) = ParseDelimitedArrayString( _
                SourceString, _
                ValueSeparator, _
                opener, _
                closer, _
                cursor _
            )
            i = i + 1
        End If
        currentChar = Mid$(SourceString, cursor, 1)
        Do Until currentChar = opener
            cursor = cursor + 1
            currentChar = Mid$(SourceString, cursor, 1)
            If currentChar = closer Or cursor >= Len(SourceString) Then
                breakLoop = True
                Exit Do
            End If
        Loop
    Loop Until breakLoop
    
    ParseDelimitedArrayString = localResult
End Function

'@Description("Returns an array from a delimited string")
Private Function ParseArraySegmentFromString(ByVal SourceString As String, _
                              ByRef cursor As Long, _
                              ByVal nextCloser As Long, _
                              ByVal ValueSeparator As String) As Variant()
Attribute ParseArraySegmentFromString.VB_Description = "Returns an array from a delimited string"
    Dim segmentLength As Long
    Dim segment() As String
    segmentLength = IIf(nextCloser = 0, Len(SourceString), nextCloser - cursor)
    segment = Split(Mid$(SourceString, cursor, segmentLength), ValueSeparator)
    cursor = nextCloser
    ParseArraySegmentFromString = duckTypeStringArray(segment)
End Function

'@Description("Takes a string array. Returns a Variant array with the elements converted to an appropriate type")
Private Function duckTypeStringArray(ByRef arr() As String) As Variant()
Attribute duckTypeStringArray.VB_Description = "Takes a string array. Returns a Variant array with the elements converted to an appropriate type"
    Dim i As Long
    Dim localLowerBound As Long
    Dim localUpperBound As Long
    Dim result() As Variant
    Dim Element As String
    
    localLowerBound = LBound(arr)
    localUpperBound = UBound(arr)
    ReDim result(localLowerBound To localUpperBound)
    For i = localLowerBound To localUpperBound
        Element = arr(i)
        If UCase$(Element) = "TRUE" Or UCase$(Element) = "FALSE" Then
            result(i) = CBool(Element)
        ElseIf IsNumeric(Element) Then
            result(i) = Val(Element)
        Else
            If Left$(Element, 1) = Chr$(34) And Right$(Element, 1) = Chr$(34) Then
                'trim string delimiters
                Element = Mid$(Element, 2, Len(Element) - 2)
            End If
            result(i) = Trim$(Element)
        End If
    Next
    duckTypeStringArray = result
End Function

'@Description("Populates the ErrorDefinitions array.")
Private Sub PopulateErrorDefinitions()
Attribute PopulateErrorDefinitions.VB_Description = "Populates the ErrorDefinitions array."
    Dim Source As String
    Source = TypeName(Me)
    ErrorDefinitions(ErrorCodes.ExpectedRangeObject) = ErrorDefinitionFactory( _
        number:=ErrorCodes.ExpectedRangeObject, _
        Source:=Source, _
        description:="Range Object Expected" _
    )
    ErrorDefinitions(ErrorCodes.MaxDimensionsLimit) = ErrorDefinitionFactory( _
        number:=ErrorCodes.MaxDimensionsLimit, _
        Source:=Source, _
        description:="Cannot convert arrays with more than 20 dimensions to a multidimensional array!" _
    )
End Sub

'@Description("Returns an ErrorDefinition Type populated with the provided arguments.")
Private Function ErrorDefinitionFactory(ByVal number As Long, ByVal Source As String, ByVal description As String) As ErrorDefinition
Attribute ErrorDefinitionFactory.VB_Description = "Returns an ErrorDefinition Type populated with the provided arguments."
    Dim result As ErrorDefinition
    result.number = number
    result.Source = Source
    result.description = description
    ErrorDefinitionFactory = result
End Function

'@Description("Raises an application error based on the passed ErrorCode and the definitions defined in ErrorDefinitions.")
Private Sub RaiseError(ByVal ErrorCode As ErrorCodes, ByVal Caller As String, Optional ByVal ArgName As String)
Attribute RaiseError.VB_Description = "Raises an application error based on the passed ErrorCode and the definitions defined in ErrorDefinitions."
    Dim currentError As ErrorDefinition
    Dim localArgName As String
    If ArgName <> vbNullString Then localArgName = "Variable " & ArgName & ": "
    currentError = ErrorDefinitions(ErrorCode)
    Err.Raise currentError.number, _
              currentError.Source & "." & Caller, _
              localArgName & currentError.description
End Sub

'Used by ToExcelRange
'@Description("Ensures each element of passed array can be represented as a scalar value.")
Private Function ensureScalar1DArray(ByRef arr() As Variant) As Variant()
Attribute ensureScalar1DArray.VB_Description = "Ensures each element of passed array can be represented as a scalar value."
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
'@Description("Nests each element of a 1d array inside it's own array and returns the new jagged array.")
Private Function OneDimArrayToJagged(ByRef arr() As Variant) As Variant()
Attribute OneDimArrayToJagged.VB_Description = "Nests each element of a 1d array inside it's own array and returns the new jagged array."
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
'@Description("Converts a 1d array to a 1 column wide 2d array.")
Private Function Transpose_1DArray(ByRef arr() As Variant) As Variant()
Attribute Transpose_1DArray.VB_Description = "Converts a 1d array to a 1 column wide 2d array."
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
'@Description("Transposes the values in a 2d array. Rows become columns, columns become rows.")
Private Function Transpose_2DArray(ByRef sourceArray() As Variant) As Variant()
Attribute Transpose_2DArray.VB_Description = "Transposes the values in a 2d array. Rows become columns, columns become rows."
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
'@Description("Transposes the values in a jagged array with a depth of 2. Rows become columns, columns become rows.")
Private Function Transpose_ArrayOfArrays(ByRef sourceArray() As Variant) As Variant()
Attribute Transpose_ArrayOfArrays.VB_Description = "Transposes the values in a jagged array with a depth of 2. Rows become columns, columns become rows."
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

'@Description("Returns an empty 1d array with 1 slot.")
Private Function getEmptyArray() As Variant()
Attribute getEmptyArray.VB_Description = "Returns an empty 1d array with 1 slot."
    Dim result() As Variant
    ReDim result(pLowerBound)
    getEmptyArray = result
End Function

'@Description("Sorts in place a passed array using the QuickSort algorithm.")
Private Sub quickSort(ByRef arr() As Variant, _
                      ByVal low As Long, _
                      ByVal high As Long, _
                      ByVal ArrayType As ArrayTypes, _
                      Optional ByVal col As Long)
Attribute quickSort.VB_Description = "Sorts in place a passed array using the QuickSort algorithm."
    Dim partitionIndex As Long
    If low < high Then
        partitionIndex = qsPartition(arr, low, high, ArrayType, col)
        quickSort arr, low, partitionIndex - 1, ArrayType, col
        quickSort arr, partitionIndex + 1, high, ArrayType, col
    End If
End Sub

'@Description("Child function of quickSort. Returns the next partition index.")
Private Function qsPartition(ByRef arr() As Variant, _
                             ByVal low As Long, _
                             ByVal high As Long, _
                             ByVal ArrayType As ArrayTypes, _
                             Optional ByVal col As Long) As Long
Attribute qsPartition.VB_Description = "Child function of quickSort. Returns the next partition index."
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

'@Description("Swaps the location of two elements in an array.")
Private Sub Swap(ByRef arr() As Variant, ByVal i As Long, ByVal j As Long)
Attribute Swap.VB_Description = "Swaps the location of two elements in an array."
    Dim Element As Variant
    letOrSetElement Element, arr(i)
    letOrSetElement arr(i), arr(j)
    '@Ignore UnassignedVariableUsage
    letOrSetElement arr(j), Element
End Sub

'@Description("Sets or lets the value of a variable by reference depending on the type of the element to be assigned")
Private Sub letOrSetElement(ByRef Destination As Variant, ByRef Source As Variant)
Attribute letOrSetElement.VB_Description = "Sets or lets the value of a variable by reference depending on the type of the element to be assigned"
    If IsObject(Source) Then
        Set Destination = Source
    Else
        Destination = Source
    End If
End Sub

'@Description("Retrieves the stored items trimmed to the correct length. Does not change structure (i.e jagged to multi).")
Private Function getInternalItems() As Variant()
Attribute getInternalItems.VB_Description = "Retrieves the stored items trimmed to the correct length. Does not change structure (i.e jagged to multi)."
    Dim result() As Variant
    If pArrayType <> Unallocated And pArrayType <> Undefined Then
        result = pItems
        If Me.Capacity > pLength Then
            If pLength > 0 Then ReDim Preserve result(pLowerBound To Me.UpperBound)
        End If
    Else
        result = getEmptyArray()
    End If
    getInternalItems = result
End Function

'@Description("Recursively parses an array to a string representation.")
Private Function recursiveToString(ByRef arr() As Variant, _
                                   ByVal PrettyPrint As Boolean, _
                                   Optional ByVal tabs As Long = 1, _
                                   Optional ByVal DelimitStrings As Boolean) As String
Attribute recursiveToString.VB_Description = "Recursively parses an array to a string representation."
    Const tabWidth As Long = 2
    Dim result As String
    Dim sep As String
    Dim i As Long

    result = "{"
    For i = LBound(arr) To UBound(arr)
        sep = IIf(i = UBound(arr), "}", IIf(PrettyPrint, ", ", ","))
        If IsArray(arr(i)) Then
            Dim nested() As Variant
            nested = arr(i)
            If PrettyPrint Then result = result & vbCrLf & Space(tabWidth * tabs)
            result = result & recursiveToString(nested, PrettyPrint, tabs + 1, DelimitStrings)
            result = result & IIf(i = UBound(arr) And PrettyPrint, vbCrLf & Space(tabWidth * (tabs - 1)) & sep, sep)
        Else
            result = result & Replace( _
                    CStr(getScalarRepresentation(arr(i), DelimitStrings)), _
                    ",", _
                    vbNullString _
                ) & sep
        End If
    Next
    recursiveToString = result
End Function

'@Description("Returns scalar values as is or appropriate representation if not scalar.")
Private Function getScalarRepresentation(ByRef Element As Variant, _
                                         Optional ByVal DelimitStrings As Boolean) As Variant
Attribute getScalarRepresentation.VB_Description = "Returns scalar values as is or appropriate representation if not scalar."
    Dim result As Variant
    If IsObject(Element) Then
        On Error Resume Next
        'Try to retrieve default member
        result = Element
        On Error GoTo 0
        ' if there is no default member or it returns an object result = Empty
        If IsEmpty(result) Then
            result = "OBJECT"
        End If
    Else
        result = Element
    End If
    ' if the element is array or default member returns an array then convert tostring
    If IsArray(result) And Not IsEmpty(result) Then
        Dim passThruArray() As Variant
        passThruArray = result
        result = recursiveToString(passThruArray, False, DelimitStrings:=DelimitStrings)
    End If
    If Not IsNumeric(result) And DelimitStrings Then
        getScalarRepresentation = Chr$(34) & result & Chr$(34)
    Else
        getScalarRepresentation = result
    End If
    
End Function

'@Description("Returns the type of a passed array.")
Private Function getArrayType(ByVal arr As Variant) As ArrayTypes
Attribute getArrayType.VB_Description = "Returns the type of a passed array."
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

'@Description("Tests if an array has been allocated.")
Private Function isArrayAllocated(ByVal arr As Variant) As Boolean
Attribute isArrayAllocated.VB_Description = "Tests if an array has been allocated."
    On Error Resume Next
    isArrayAllocated = IsArray(arr) And _
                                    Not IsError(LBound(arr, 1)) And _
                                    LBound(arr, 1) <= UBound(arr, 1)
    On Error GoTo 0
End Function

'@Description("Tests if an array is jagged.")
Private Function isJaggedArray(ByVal arr As Variant) As Boolean
Attribute isJaggedArray.VB_Description = "Tests if an array is jagged."
    If IsArray(arr) Then
        On Error GoTo ErrHandler
        Dim Element As Variant
        For Each Element In arr
            If IsArray(Element) Then
                isJaggedArray = True
                Exit Function
            End If
        Next
        On Error GoTo 0
    End If
    Exit Function
ErrHandler:
    Err.Clear
End Function

'@Description("Tests if an array is multidimensioned.")
Private Function isMultidimensionalArray(ByVal arr As Variant) As Boolean
Attribute isMultidimensionalArray.VB_Description = "Tests if an array is multidimensioned."
    If IsArray(arr) Then
        On Error GoTo ErrHandler
        '@Ignore VariableNotUsed
        Dim localUpperBound As Long
        '@Ignore AssignmentNotUsed
        localUpperBound = UBound(arr, 2)
        isMultidimensionalArray = True
        On Error GoTo 0
    End If
    Exit Function
ErrHandler:
    Err.Clear
End Function

'@Ignore FunctionReturnValueNotUsed
'@Description("re-indexes array to conform to new bounds. Does not change length.")
Private Function Rebase(Optional ByRef recurse As Boolean, Optional ByVal arr As Variant) As Variant()
Attribute Rebase.VB_Description = "re-indexes array to conform to new bounds. Does not change length."
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

'@Description("re-indexes nested arrays to conform to new bounds. Child function of Rebase)
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

'@Description("Increases capacity of internal array if required.")
Private Sub EnsureCapacity(ByVal MinimumCapacity As Long)
Attribute EnsureCapacity.VB_Description = "Increases capacity of internal array if required."
    If Me.Capacity < MinimumCapacity Then
        Dim newCapacity As Long
        newCapacity = IIf(Me.Capacity = 0, DEFAULT_CAPACITY, Me.Capacity * 2)
            If newCapacity > MAX_ARRAY_LENGTH Then newCapacity = MAX_ARRAY_LENGTH
        If newCapacity < MinimumCapacity Then newCapacity = MinimumCapacity
        Me.Capacity = newCapacity
    End If
End Sub

'@Description("Fills the newItems arg by ref with the values stored in the internal array at the corresponding index.")
Private Function copyRange(ByRef newItems() As Variant) As Variant()
Attribute copyRange.VB_Description = "Fills the newItems arg by ref with the values stored in the internal array at the corresponding index."
    Dim i As Long
    For i = LBound(newItems) To UBound(newItems)
        newItems(i) = pItems(i)
    Next
    copyRange = newItems
End Function

'@Description("Converts jagged arrays to multidimension arrays")
Private Function jaggedToMulti(ByRef arr() As Variant, _
                      Optional ByVal depth As Long, _
                      Optional ByVal currentDepth As Long, _
                      Optional ByRef crumbs As Variant, _
                      Optional ByRef result As Variant, _
                      Optional ByVal EnsureScalar As Boolean) As Variant()
Attribute jaggedToMulti.VB_Description = "Converts jagged arrays to multidimension arrays"
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
            ' TODO: raise custom err
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

'@Description("Converts multidimension arrays to jagged arrays.")
Private Function multiToJagged(ByRef arr() As Variant, _
                      Optional ByVal depth As Long, _
                      Optional ByVal currentDepth As Long, _
                      Optional ByRef crumbs As Variant) As Variant()
Attribute multiToJagged.VB_Description = "Converts multidimension arrays to jagged arrays."
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
            ' TODO: raise custom err
            Debug.Print "Argument passed to 'multiToJagged' not a multidimensional array"
            Exit Function
        End If
        localDepth = getMultidimensionalArrayDepth(arr)
    End If
    
    If localDepth > 20 Then
        ' too many dimensions to be handled by getElementByBreadcrumb
        ' handle error
        ' TODO: raise custom err
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

'@Description("Returns a number representing how many levels of nested arrays are contained in array.")
Private Function getJaggedArrayDepth(ByRef arr() As Variant) As Long
Attribute getJaggedArrayDepth.VB_Description = "Returns a number representing how many levels of nested arrays are contained in array."
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

'@Description("Gets the maximium bound values for a specified dimension of a jagged array.")
Private Function getMaxBoundsAtDimension(ByRef sourceArray() As Variant, _
                                         Optional ByVal Dimension As Long = 1, _
                                         Optional ByVal currentDimension As Long = 1) As Variant()
Attribute getMaxBoundsAtDimension.VB_Description = "Gets the maximium bound values for a specified dimension of a jagged array."
    Dim maxBounds() As Variant
    
    If currentDimension < Dimension Then
        Dim i As Long
        Dim currentBounds() As Variant
        Dim localLowerBound As Long
        Dim localUpperBound As Long
        
        localLowerBound = LBound(sourceArray)
        localUpperBound = UBound(sourceArray)
        For i = localLowerBound To localUpperBound
            If IsArray(sourceArray(i)) Then
                Dim nested() As Variant
                nested = sourceArray(i)
                currentBounds = getMaxBoundsAtDimension(nested, Dimension, currentDimension + 1)
                If Not isArrayAllocated(maxBounds) Then
                    maxBounds = currentBounds
                Else
                    If currentBounds(0) < maxBounds(0) Then maxBounds(0) = currentBounds(0)
                    If currentBounds(1) > maxBounds(1) Then maxBounds(1) = currentBounds(1)
                End If
            End If
        Next
    Else
        maxBounds = getArrayBounds(sourceArray)
    End If
    
    getMaxBoundsAtDimension = maxBounds
End Function

'@Description("Returns an array representing the structure of a jagged array")
Private Function mapJaggedArray(ByRef sourceArray() As Variant, _
                       Optional ByVal knownDepth As Long) As Variant()
Attribute mapJaggedArray.VB_Description = "Returns an array representing the structure of a jagged array"
    Dim i As Long
    Dim localMap() As Variant
    Dim localDepth As Long

    If Not IsArray(sourceArray) Then Exit Function

    If knownDepth > 0 Then
        localDepth = knownDepth
    Else
        localDepth = getJaggedArrayDepth(sourceArray)
    End If

    ' crumbs are base 0
    ReDim localMap(0 To localDepth - 1)

    For i = LBound(localMap) To UBound(localMap)
        localMap(i) = getMaxBoundsAtDimension(sourceArray, i + 1)
    Next

    mapJaggedArray = localMap
End Function

'@Description("Returns a 0-based array with two slots. First slot is the lowerbound of the passed array. second slot (index 0) is the upper bound.")
Private Function getArrayBounds(ByVal arr As Variant) As Variant()
Attribute getArrayBounds.VB_Description = "Returns a 0-based array with two slots. First slot is the lowerbound of the passed array. second slot (index 0) is the upper bound."
    Dim result(0 To 1) As Variant
    If IsArray(arr) Then
        result(0) = LBound(arr)
        result(1) = UBound(arr)
    End If
    getArrayBounds = result
End Function

'@Description("Returns a number representing the total number of dimensions of a multidimension array.")
Private Function getMultidimensionalArrayDepth(ByVal arr As Variant) As Long
Attribute getMultidimensionalArrayDepth.VB_Description = "Returns a number representing the total number of dimensions of a multidimension array."
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

'@Description("Returns the element stored in an array based on a crumb argument. A crumb is an array representing the location of the multidimension array slot.")
Private Function getElementByBreadcrumb(ByRef arr() As Variant, ByRef crumb() As Variant) As Variant
Attribute getElementByBreadcrumb.VB_Description = "Returns the element stored in an array based on a crumb argument. A crumb is an array representing the location of the multidimension array slot."
    Dim result As Variant
    Select Case UBound(crumb)
    Case 0
        letOrSetElement result, arr(crumb(0))
    Case 1
        letOrSetElement result, arr(crumb(0), crumb(1))
    Case 2
        letOrSetElement result, arr(crumb(0), crumb(1), crumb(2))
    Case 3
        letOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3))
    Case 4
        letOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4))
    Case 5
        letOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5))
    Case 6
        letOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6))
    Case 7
        letOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7))
    Case 8
        letOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8))
    Case 9
        letOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9))
    Case 10
        letOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10))
    Case 11
        letOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10), crumb(11))
    Case 12
        letOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10), crumb(11), crumb(12))
    Case 13
        letOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10), crumb(11), crumb(12), crumb(13))
    Case 14
        letOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10), crumb(11), crumb(12), crumb(13), crumb(14))
    Case 15
        letOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15))
    Case 16
        letOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16))
    Case 17
        letOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17))
    Case 18
        letOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18))
    Case 19
        letOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18), crumb(19))
    Case 20
        letOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18), crumb(19), crumb(20))
    End Select
    '@Ignore UnassignedVariableUsage
    getElementByBreadcrumb = result
End Function

'@Description("Assigns an element to a location in a multidimensonal array based on a crumb argument. A crumb is an array representing the location of the multidimension array slot.")
Private Function letElementByBreadcrumb(ByRef arr() As Variant, ByRef crumb() As Variant, ByVal Element As Variant) As Variant
Attribute letElementByBreadcrumb.VB_Description = "Assigns an element to a location in a multidimensonal array based on a crumb argument. A crumb is an array representing the location of the multidimension array slot."
    Select Case UBound(crumb)
    Case 0
        letOrSetElement arr(crumb(0)), Element
    Case 1
        letOrSetElement arr(crumb(0), crumb(1)), Element
    Case 2
        letOrSetElement arr(crumb(0), crumb(1), crumb(2)), Element
    Case 3
        letOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3)), Element
    Case 4
        letOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4)), Element
    Case 5
        letOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5)), Element
    Case 6
        letOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6)), Element
    Case 7
        letOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7)), Element
    Case 8
        letOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8)), Element
    Case 9
        letOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9)), _
            Element
    Case 10
        letOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
            crumb(10)), Element
    Case 11
        letOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
            crumb(10), crumb(11)), Element
    Case 12
        letOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
            crumb(10), crumb(11), crumb(12)), Element
    Case 13
        letOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
            crumb(10), crumb(11), crumb(12), crumb(13)), Element
    Case 14
        letOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
            crumb(10), crumb(11), crumb(12), crumb(13), crumb(14)), Element
    Case 15
        letOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
            crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15)), Element
    Case 16
        letOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
        crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16)), Element
    Case 17
        letOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
            crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17)), Element
    Case 18
        letOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
            crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18)), Element
    Case 19
        letOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
        crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18), crumb(19)), Element
    Case 20
        letOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
        crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18), crumb(19), crumb(20)), _
        Element
    End Select

    letElementByBreadcrumb = arr
End Function

'@Description("Returns an empty multidimensional array with the dimensions specified in the crumb argument. A crumb is an array representing the structure of the multidimension array.")
Private Function createMultidimensionalArray(ByRef crumb() As Variant) As Variant()
Attribute createMultidimensionalArray.VB_Description = "Returns an empty multidimensional array with the dimensions specified in the crumb argument. A crumb is an array representing the structure of the multidimension array."
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













