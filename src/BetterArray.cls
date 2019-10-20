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
'@Folder("BetterArray")
'@ModuleDescription("An array class for VBA providing features found in more modern languages ")

Option Explicit

'''''''''''''''''''
' Constant Fields '
'''''''''''''''''''

Private Const DEFAULT_CAPACITY As Long = 4
Private Const MAX_ARRAY_LENGTH = &H7FEFFFFF
Private Enum ArrayTypes
    BA_UNDEFINED
    BA_UNALLOCATED
    BA_ONEDIMENSION
    BA_MULTIDIMENSION
    BA_JAGGED
End Enum

Private Enum ErrorCodes
    EC_START = vbObjectError + 512
    EC_EXPECTED_RANGE_OBJECT
    EC_EXPECTED_COLLECTION_OBJECT
    EC_MAX_DIMENSIONS_LIMIT
    EC_EXCEEDS_MAX_SORT_DEPTH
    EC_EXPECTED_JAGGED_ARRAY
    EC_EXPECTED_MULTIDIMENSION_ARRAY
    EC_EXPECTED_ARRAY
    EC_END
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
Private ErrorDefinitions(EC_START To EC_END) As ErrorDefinition

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
Public Property Get Item(ByVal Index As Long)
Attribute Item.VB_Description = "Returns the element in the array stored at the specified index."
Attribute Item.VB_UserMemId = 0
    If Index <= pLength Then
        ' not using letOrSetElement here as it might try to access the Items letter
        If IsObject(pItems(Index)) Then
            Set Item = pItems(Index)
        Else
            Item = pItems(Index)
        End If
    Else
        Err.Raise 9
    End If
End Property

'@Description("Assigns the passed element to the array at the specified index.")
Public Property Let Item(ByVal Index As Long, ByVal Element As Variant)
Attribute Item.VB_Description = "Assigns the passed element to the array at the specified index."
    If Me.UpperBound >= Index Then
        If Index < Me.LowerBound Then
            Me.Unshift Element
        Else
            letOrSetElement pItems(Index), Element
        End If
    Else
        Me.Push Element
    End If
End Property

'@Description("Returns the stored array.")
Public Property Get Items() As Variant
Attribute Items.VB_Description = "Returns the stored array."
    Dim Result() As Variant
    Result = getInternalItems()
    If pArrayType = ArrayTypes.BA_MULTIDIMENSION And isJaggedArray(Result) Then Result = jaggedToMulti(Result)
    Items = Result
End Property

'@Description("Stores the passed array.")
Public Property Let Items(ByVal Values As Variant)
Attribute Items.VB_Description = "Stores the passed array."
    Dim localLowerBound As Long
    Dim localValues() As Variant
    If IsArray(Values) Then
        localValues = Values
        pArrayType = getArrayType(localValues)
        If pArrayType = ArrayTypes.BA_UNDEFINED _
           Or pArrayType = ArrayTypes.BA_UNALLOCATED Then
            ' TOD: Define custom error
            Err.Raise 9
        Else
            If pArrayType = ArrayTypes.BA_MULTIDIMENSION Then
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
Public Function Push(ParamArray Args() As Variant) As Variant
Attribute Push.VB_Description = "Adds the provided argments to the array at the next available index."
    Dim Element As Variant

    If pArrayType = ArrayTypes.BA_UNALLOCATED Or _
       pArrayType = ArrayTypes.BA_UNDEFINED Then
        pArrayType = ArrayTypes.BA_ONEDIMENSION
    End If

    For Each Element In Args
        If IsArray(Element) And pArrayType = ArrayTypes.BA_ONEDIMENSION Then
            pArrayType = ArrayTypes.BA_JAGGED
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
    Dim Result As Variant

    If pLength > 0 Then
        Result = pItems(Me.UpperBound)
        pLength = pLength - 1
        ReDim newItems(pLowerBound To Me.UpperBound)
        Me.Items = copyRange(newItems)
    End If
    Pop = Result
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Removes the first element from the array and returns that element.")
Public Function Shift() As Variant
Attribute Shift.VB_Description = "Removes the first element from the array and returns that element."
    Dim Result As Variant
    Dim newItems() As Variant
    If pLength > 0 Then
        Result = pItems(pLowerBound)
        ReDim newItems((pLowerBound + 1) To Me.UpperBound)
        Me.Items = copyRange(newItems)
    End If
    Shift = Result
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Adds one or more elements to the beginning of an array. Returns the new length of the array.")
Public Function Unshift(ParamArray Args() As Variant) As Long
Attribute Unshift.VB_Description = "Adds one or more elements to the beginning of an array. Returns the new length of the array."
    Dim newItems() As Variant
    Dim oldItems() As Variant
    Dim oldType As ArrayTypes
    
    newItems = Args
    oldType = pArrayType
    oldItems = getInternalItems()
    Me.Items = newItems
    If oldType <> BA_UNALLOCATED And oldType <> BA_UNDEFINED Then
        Me.Concat oldItems
        pArrayType = oldType
    End If
    Unshift = Me.Length
End Function

' TODO: Document
'@Ignore FunctionReturnValueNotUsed
'@Description("Parses a string representing an array and stores the outcome.")
Public Function ParseFromString( _
        ByVal SourceString As String, _
        Optional ByVal ValueSeparator As String = ",", _
        Optional ByVal ArrayOpenDelimiter As String, _
        Optional ByVal ArrayClosingDelimiter As String _
    ) As BetterArray
Attribute ParseFromString.VB_Description = "Parses a string representing an array and stores the outcome."
    Dim Opener As String
    Dim Closer As String
    Dim ArraysAreDelimited As Boolean
    Dim Result() As Variant
    
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
                Opener = firstChar
                Closer = lastChar
                ArraysAreDelimited = True
            End If
        Else
            ' array delimiters supplied
            Opener = ArrayOpenDelimiter
            Closer = ArrayClosingDelimiter
            ArraysAreDelimited = True
        End If
        
        If ArraysAreDelimited Then
            Result = ParseDelimitedArrayString(SourceString, ValueSeparator, Opener, Closer)
        Else
            Result = ParseArraySegmentFromString(SourceString, 1, 0, ValueSeparator)
        End If
    Else
        ' nullstring passed
        ' handle exception
    End If
    
    Me.Items = Result
    Set ParseFromString = Me
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Returns a string representing the array structure and its elements.")
Public Function ToString( _
        Optional ByVal PrettyPrint As Boolean, _
        Optional ByVal DelimitStrings As Boolean _
    ) As String
Attribute ToString.VB_Description = "Returns a string representing the array structure and its elements."
    Dim ArrayType As ArrayTypes
    Dim Result As String
    Dim localItems() As Variant
    
    localItems = getInternalItems()
    ArrayType = getArrayType(localItems)
    If Not ArrayType = ArrayTypes.BA_UNDEFINED And _
       Not ArrayType = ArrayTypes.BA_UNALLOCATED Then
        If ArrayType = ArrayTypes.BA_MULTIDIMENSION Then
            localItems = multiToJagged(localItems)
        End If
        Result = recursiveToString(localItems, PrettyPrint, DelimitStrings:=DelimitStrings)
    End If
    ToString = Result
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Determines whether the array includes a certain value among its entries.")
Public Function Includes( _
        ByVal SearchElement As Variant, _
        Optional ByVal FromIndex As Long _
    ) As Boolean
Attribute Includes.VB_Description = "Determines whether the array includes a certain value among its entries."
    Dim localLength As Long
    Dim currentIndex As Long
    Dim Arr() As Variant
    
    Arr = getInternalItems()
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
        If Arr(currentIndex) = SearchElement Then
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
    Dim Result() As Variant
    
    localLowerBound = Me.LowerBound
    localUpperBound = Me.UpperBound
    ReDim Result(localLowerBound To localUpperBound)
    For i = localLowerBound To localUpperBound: Result(i) = i: Next
    Keys = Result
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Returns the largest value in a list of values. If arguments are provided it will return the largest argument. If no arguments are provided it will return the largest element in the stored array")
Public Function Max(ParamArray Args() As Variant) As Variant
Attribute Max.VB_Description = "Returns the largest value in a list of values. If arguments are provided it will return the largest argument. If no arguments are provided it will return the largest element in the stored array"
    Dim i As Long
    Dim Arr() As Variant
    Dim Result As Variant
    
    If UBound(Args) < LBound(Args) Then          'Test if no args provided
        Arr = getInternalItems()
    Else
        Arr = Args
        If IsArray(Arr) Then
            ' exit function if array is multi dimensional
            If isMultidimensionalArray(Arr) Then
                Exit Function
            End If
        End If
    End If
    
    ' extract first array if argument(s) are arrays
    Do While IsArray(Arr(0))
        Arr = Arr(0)
    Loop
    
    For i = LBound(Arr) To UBound(Arr)
        If Not IsArray(Arr(i)) And Not IsObject(Arr(i)) Then
            If IsEmpty(Result) Then
                Result = Arr(i)
            Else
                If Arr(i) > Result Then Result = Arr(i)
            End If
        End If
    Next
    
    Max = Result
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Returns the smallest value in a list of values. If arguments are provided it will return the smallest argument. If no arguments are provided it will return the smallest element in the stored array")
Public Function Min(ParamArray Args() As Variant) As Variant
Attribute Min.VB_Description = "Returns the smallest value in a list of values. If arguments are provided it will return the smallest argument. If no arguments are provided it will return the smallest element in the stored array"
    Dim i As Long
    Dim Arr() As Variant
    Dim Result As Variant
    
    If UBound(Args) < LBound(Args) Then          'Test if no args provided
        Arr = getInternalItems()
    Else
        Arr = Args
        If IsArray(Arr) Then
            ' exit function if array is multi dimensional
            If isMultidimensionalArray(Arr) Then
                Exit Function
            End If
        End If
    End If
    
    ' extract first array if argument(s) are arrays
    Do While IsArray(Arr(0))
        Arr = Arr(0)
    Loop
        
    For i = LBound(Arr) To UBound(Arr)
        If Not IsArray(Arr(i)) And Not IsObject(Arr(i)) Then
            If IsEmpty(Result) Then
                Result = Arr(i)
            Else
                If Arr(i) > Result Then Result = Arr(i)
            End If
        End If
    Next
    
    Min = Result
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
    Dim Arr() As Variant
    Dim Result() As Variant
    
    Arr = getInternalItems()
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
    ReDim Result(newIndex To count - 1)
    Do While oldIndex < final
        If oldIndex >= LBound(Arr) And oldIndex <= UBound(Arr) Then
            Result = Arr(oldIndex)
            newIndex = newIndex + 1
            oldIndex = oldIndex + 1
        End If
    Loop
    
    Slice = Result
End Function

' TODO: Implement, Document & Test FromExcelRange method
'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Stores the values contained in an Excel range in the internal array")
Public Function FromExcelRange( _
        ByRef FromRange As Object, _
        Optional ByVal ColumnNumber As Long, _
        Optional ByVal RowNumber As Long _
    ) As BetterArray
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
        RaiseError ErrorCodes.EC_EXPECTED_RANGE_OBJECT, "FromExcelRange()", "FromRange"
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
    Dim SourceArray() As Variant
    Dim Result() As Variant
    
    SourceArray = getInternalItems()
    If pArrayType = ArrayTypes.BA_MULTIDIMENSION Or pArrayType = ArrayTypes.BA_JAGGED Then
        localLowerBound = LBound(SourceArray)
        localUpperBound = UBound(SourceArray)
        ReDim Result(localLowerBound To localUpperBound)
        
        If ColumnNumber >= localLowerBound And ColumnNumber <= localUpperBound Then
            localColumnNumber = ColumnNumber
        Else
            localColumnNumber = Me.LowerBound
        End If
        
        For i = localLowerBound To localUpperBound
            letOrSetElement Result(i), SourceArray(i)(localColumnNumber)
        Next
    Else
        ' can't extract from 1d or empty - return unchanged list
        Result = SourceArray
    End If
    
    ExtractColumn = Result
End Function

'@Description("Extracts the specified row of a multidimensional/jagged array as a new array.")
Public Function ExtractRow(Optional ByVal RowNumber As Long) As Variant()
Attribute ExtractRow.VB_Description = "Extracts the specified row of a multidimensional/jagged array as a new array."
    Dim localRowNumber As Long
    Dim localLowerBound As Long
    Dim localUpperBound As Long
    Dim SourceArray() As Variant
    Dim Result() As Variant

    SourceArray = getInternalItems()
    If pArrayType = ArrayTypes.BA_MULTIDIMENSION Or pArrayType = ArrayTypes.BA_JAGGED Then
        localLowerBound = LBound(SourceArray)
        localUpperBound = UBound(SourceArray)
        ReDim Result(localLowerBound To localUpperBound)

        If RowNumber >= localLowerBound And RowNumber <= localUpperBound Then
            localRowNumber = RowNumber
        Else
            localRowNumber = Me.LowerBound
        End If
        letOrSetElement Result, SourceArray(localRowNumber)
    Else
        ' can't extract from 1d or empty - return
        Result = SourceArray
    End If
    ExtractRow = Result
End Function

' TODO: Implement, Document & Test ToExcelRange method
'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Writes the values stored in the array to an excel worksheet starting at the specified range.")
Public Function ToExcelRange( _
        ByRef Destination As Object, _
        Optional ByVal TransposeValues As Boolean _
    ) As Object
Attribute ToExcelRange.VB_Description = "Writes the values stored in the array to an excel worksheet starting at the specified range."
    Dim localRange As Object
    Dim localItems() As Variant
    Dim Depth As Long
    Dim lengthRows As Long
    Dim lengthColumns As Long
    
    lengthRows = 1 'default
    lengthColumns = 1 'default
    
    If TypeName(Destination) = "Range" Then
        localItems = getInternalItems()
        Depth = getJaggedArrayDepth(localItems)
        If Depth > 0 Then 'check array is allocated
            If Depth = 1 Then
                '1d array
                localItems = ensureScalar1DArray(localItems)
                If TransposeValues Then
                    lengthRows = Me.Length
                    localItems = Transpose_1DArray(localItems)
                Else
                    lengthColumns = Me.Length
                End If
            ElseIf Depth > 1 Then
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
        RaiseError ErrorCodes.EC_EXPECTED_RANGE_OBJECT, "ToExcelRange()", "Destination"
    End If
    
    Set ToExcelRange = localRange
End Function

''''''''''''''''''''''''''''
' Chainable Public Methods '
''''''''''''''''''''''''''''
' TODO: Document & Test
'@Ignore ProcedureNotUsed
Public Function Transpose() As BetterArray
    Dim Result() As Variant
    Dim localItems() As Variant
    Dim localType As ArrayTypes
    localItems = getInternalItems()
    localType = getArrayType(localItems)
    Select Case localType
    Case ArrayTypes.BA_ONEDIMENSION
        Result = Transpose_1DArray(localItems)
    Case ArrayTypes.BA_MULTIDIMENSION
        ' Sould Never match as MD arrs stored as jagged internally
        Result = Transpose_2DArray(localItems)
    Case ArrayTypes.BA_JAGGED
        Result = Transpose_ArrayOfArrays(localItems)
    Case Else
        Result = localItems
    End Select
    Me.Items = Result
    Set Transpose = Me
End Function


' TODO: Document & Test - Clone
'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Returns a new BetterArray instance containing the same values as the current instance.")
Public Function Clone() As BetterArray
Attribute Clone.VB_Description = "Returns a new BetterArray instance containing the same values as the current instance."
    Dim Result As BetterArray
    Set Result = New BetterArray
    Result.LowerBound = Me.LowerBound
    Result.Items = Me.Items
    Set Clone = Result
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
Public Function Concat(ParamArray Args() As Variant) As BetterArray
Attribute Concat.VB_Description = "Joins one or more arrays onto the end of the current array."
    Dim arg As Variant
    Dim localArray() As Variant
    Dim localType As ArrayTypes
    Dim i As Long
    
    
    For Each arg In Args
        If IsArray(arg) Then
            localArray = arg
            localType = getArrayType(localArray)
            Select Case pArrayType
                Case BA_UNALLOCATED, BA_UNDEFINED, BA_ONEDIMENSION
                    pArrayType = localType
            End Select
            If localType = BA_MULTIDIMENSION Then
                localArray = multiToJagged(localArray)
            End If
            For i = LBound(localArray) To UBound(localArray)
                Me.Push localArray(i)
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
    If SourceCollection Is Nothing Then
        RaiseError EC_EXPECTED_COLLECTION_OBJECT, "CopyFromCollection", "SourceCollection"
    End If
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
    If ArrayType = ArrayTypes.BA_UNALLOCATED Or _
       ArrayType = ArrayTypes.BA_UNDEFINED Then
        sortedItems = getEmptyArray
    Else
        If Me.Length > 0 Then
            If ArrayType <> BA_ONEDIMENSION Then
                If ArrayType = ArrayTypes.BA_MULTIDIMENSION Then
                    localItems = multiToJagged(localItems)
                End If
                Dim Depth As Long
                Depth = getJaggedArrayDepth(localItems)
                If Depth > 2 Then
                    RaiseError EC_EXCEEDS_MAX_SORT_DEPTH, "Sort"
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
Public Function CopyWithin( _
        ByVal Target As Long, _
        Optional ByVal StartIndex As Long, _
        Optional ByVal EndIndex As Long _
    ) As BetterArray
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
    Dim Arr() As Variant
    
    Arr = getInternalItems()
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
        If from >= LBound(Arr) And from <= UBound(Arr) Then
            Arr(idx) = Arr(from)
        End If
        from = from + direction
        idx = idx + direction
        count = count - 1
    Loop
    
    Me.Items = Arr
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
    Dim Arr() As Variant
    Dim Result() As Variant
    Arr = getInternalItems()
    localUpperBound = Me.UpperBound
    ReDim Result(Me.LowerBound To localUpperBound)
    For i = localUpperBound To Me.LowerBound Step -1
        Result = Arr(i)
    Next
    Me.Items = Result
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
Private Function ParseDelimitedArrayString( _
        ByVal SourceString As String, _
        ByVal ValueSeparator As String, _
        ByVal Opener As String, _
        ByVal Closer As String, _
        Optional ByRef Cursor As Long = 2 _
    ) As Variant()
Attribute ParseDelimitedArrayString.VB_Description = "Recursively parses nested arrays from a string."
    Dim currentChar As String
    Dim localResult() As Variant
    Dim i As Long
    Dim breakLoop As Boolean
    Dim nextOpener As Long
    Dim NextCloser As Long
    
    currentChar = Mid$(SourceString, Cursor, 1)
    nextOpener = InStr(Cursor, SourceString, Opener)
    NextCloser = InStr(Cursor, SourceString, Closer)
    
    If currentChar <> Opener And (NextCloser < nextOpener Or nextOpener = 0) Then
        ' we're inside an array to be parsed
        ParseDelimitedArrayString = ParseArraySegmentFromString(SourceString, Cursor, NextCloser, ValueSeparator)
        Exit Function
    End If
        
    ' traverse nested arrays
    i = pLowerBound
    Do
        If currentChar = Opener Then
            If i > pLowerBound Then
                ReDim Preserve localResult(i)
            Else
                localResult = getEmptyArray
            End If
            Cursor = Cursor + 1
            localResult(i) = ParseDelimitedArrayString( _
                SourceString, _
                ValueSeparator, _
                Opener, _
                Closer, _
                Cursor _
            )
            i = i + 1
        End If
        currentChar = Mid$(SourceString, Cursor, 1)
        Do Until currentChar = Opener
            Cursor = Cursor + 1
            currentChar = Mid$(SourceString, Cursor, 1)
            If currentChar = Closer Or Cursor >= Len(SourceString) Then
                breakLoop = True
                Exit Do
            End If
        Loop
    Loop Until breakLoop
    
    ParseDelimitedArrayString = localResult
End Function

'@Description("Returns an array from a delimited string")
Private Function ParseArraySegmentFromString( _
        ByVal SourceString As String, _
        ByRef Cursor As Long, _
        ByVal NextCloser As Long, _
        ByVal ValueSeparator As String _
    ) As Variant()
Attribute ParseArraySegmentFromString.VB_Description = "Returns an array from a delimited string"
    Dim segmentLength As Long
    Dim segment() As String
    segmentLength = IIf(NextCloser = 0, Len(SourceString), NextCloser - Cursor)
    segment = Split(Mid$(SourceString, Cursor, segmentLength), ValueSeparator)
    Cursor = NextCloser
    ParseArraySegmentFromString = duckTypeStringArray(segment)
End Function

'@Description("Takes a string array. Returns a Variant array with the elements converted to an appropriate type")
Private Function duckTypeStringArray(ByRef Arr() As String) As Variant()
Attribute duckTypeStringArray.VB_Description = "Takes a string array. Returns a Variant array with the elements converted to an appropriate type"
    Dim i As Long
    Dim localLowerBound As Long
    Dim localUpperBound As Long
    Dim Result() As Variant
    Dim Element As String
    
    localLowerBound = LBound(Arr)
    localUpperBound = UBound(Arr)
    ReDim Result(localLowerBound To localUpperBound)
    For i = localLowerBound To localUpperBound
        Element = Arr(i)
        If UCase$(Element) = "TRUE" Or UCase$(Element) = "FALSE" Then
            Result(i) = CBool(Element)
        ElseIf IsNumeric(Element) Then
            Const Epsilon As Double = 2 ^ -52
            Dim diff As Double
            ' check if integer
            diff = Abs(Fix(Element) - Element)
            If diff > Epsilon Then
                Result(i) = CDbl(Element)
            Else
                Result(i) = CLng(Element)
            End If
        Else
            If Left$(Element, 1) = Chr$(34) And Right$(Element, 1) = Chr$(34) Then
                'trim string delimiters
                Element = Mid$(Element, 2, Len(Element) - 2)
            End If
            Result(i) = Trim$(Element)
        End If
    Next
    duckTypeStringArray = Result
End Function

'@Description("Populates the ErrorDefinitions array.")
Private Sub PopulateErrorDefinitions()
Attribute PopulateErrorDefinitions.VB_Description = "Populates the ErrorDefinitions array."
    Dim Source As String
    Source = TypeName(Me)
    ErrorDefinitions(ErrorCodes.EC_EXPECTED_RANGE_OBJECT) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_EXPECTED_RANGE_OBJECT, _
        Source:=Source, _
        description:="Range Object Expected" _
    )
    ErrorDefinitions(ErrorCodes.EC_MAX_DIMENSIONS_LIMIT) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_MAX_DIMENSIONS_LIMIT, _
        Source:=Source, _
        description:="Cannot convert structure of arrays with more than 20 dimensions." _
    )
    ErrorDefinitions(ErrorCodes.EC_EXPECTED_COLLECTION_OBJECT) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_EXPECTED_COLLECTION_OBJECT, _
        Source:=Source, _
        description:="Valid Collection Object Expected" _
    )
    ErrorDefinitions(ErrorCodes.EC_EXCEEDS_MAX_SORT_DEPTH) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_EXCEEDS_MAX_SORT_DEPTH, _
        Source:=Source, _
        description:="Cannot sort on arrays with more than 2 dimensions" _
    )
    ErrorDefinitions(ErrorCodes.EC_EXPECTED_JAGGED_ARRAY) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_EXPECTED_JAGGED_ARRAY, _
        Source:=Source, _
        description:="Expected jagged array." _
    )
    ErrorDefinitions(ErrorCodes.EC_EXPECTED_MULTIDIMENSION_ARRAY) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_EXPECTED_MULTIDIMENSION_ARRAY, _
        Source:=Source, _
        description:="Expected multidimension array." _
    )
    ErrorDefinitions(ErrorCodes.EC_EXPECTED_ARRAY) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_EXPECTED_ARRAY, _
        Source:=Source, _
        description:="Expected array." _
    )
    
End Sub

'@Description("Returns an ErrorDefinition Type populated with the provided arguments.")
Private Function ErrorDefinitionFactory( _
        ByVal number As Long, _
        ByVal Source As String, _
        ByVal description As String _
    ) As ErrorDefinition
Attribute ErrorDefinitionFactory.VB_Description = "Returns an ErrorDefinition Type populated with the provided arguments."
    Dim Result As ErrorDefinition
    Result.number = number
    Result.Source = Source
    Result.description = description
    ErrorDefinitionFactory = Result
End Function

'@Description("Raises an application error based on the passed ErrorCode and the definitions defined in ErrorDefinitions.")
Private Sub RaiseError( _
        ByVal ErrorCode As ErrorCodes, _
        ByVal Caller As String, _
        Optional ByVal ArgName As String _
    )
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
Private Function ensureScalar1DArray(ByRef Arr() As Variant) As Variant()
Attribute ensureScalar1DArray.VB_Description = "Ensures each element of passed array can be represented as a scalar value."
    Dim i As Long
    Dim localLowerBound As Long
    Dim localUpperBound As Long
    Dim Result() As Variant
    
    localLowerBound = LBound(Arr)
    localUpperBound = UBound(Arr)
    ReDim Result(localLowerBound To localUpperBound)
    For i = localLowerBound To localUpperBound
        Result(i) = getScalarRepresentation(Arr(i))
    Next
    ensureScalar1DArray = Result
End Function

'@Description("Converts a 1d array to a 1 column wide 2d array.")
Private Function Transpose_1DArray(ByRef Arr() As Variant) As Variant()
Attribute Transpose_1DArray.VB_Description = "Converts a 1d array to a 1 column wide 2d array."
    'TODO: need check to ensure 1d?
    Dim i As Long
    Dim j As Long
    Dim localUpperBound As Long
    Dim localLowerBound As Long
    Dim Result() As Variant
    localUpperBound = UBound(Arr)
    localLowerBound = LBound(Arr)
    ReDim Result(localLowerBound To (localUpperBound - localLowerBound + 1), _
                 localLowerBound To localLowerBound)
    j = localLowerBound
    For i = localLowerBound To localUpperBound
        Result(j, localLowerBound) = Arr(i)
        j = j + 1
    Next
    Transpose_1DArray = Result
End Function

'@Description("Transposes the values in a 2d array. Rows become columns, columns become rows.")
Private Function Transpose_2DArray(ByRef SourceArray() As Variant) As Variant()
Attribute Transpose_2DArray.VB_Description = "Transposes the values in a 2d array. Rows become columns, columns become rows."
    Dim currentRow As Long
    Dim LowerBoundRow As Long
    Dim UpperBoundRow As Long
    Dim currentColumn As Long
    Dim LowerBoundCol As Long
    Dim UpperBoundCol As Long
    Dim Result() As Variant
    
    LowerBoundCol = LBound(SourceArray, 1)
    UpperBoundCol = UBound(SourceArray, 1)
    LowerBoundRow = LBound(SourceArray, 2)
    UpperBoundRow = UBound(SourceArray, 2)
    ReDim Result(LowerBoundRow To UpperBoundRow, LowerBoundCol To UpperBoundCol)
    For currentRow = LowerBoundRow To UpperBoundRow
        For currentColumn = LowerBoundCol To UpperBoundRow
            Result(currentRow, currentColumn) = SourceArray(currentColumn, currentRow)
        Next
    Next
    Transpose_2DArray = Result
End Function

'@Description("Transposes the values in a jagged array with a depth of 2. Rows become columns, columns become rows.")
Private Function Transpose_ArrayOfArrays(ByRef SourceArray() As Variant) As Variant()
Attribute Transpose_ArrayOfArrays.VB_Description = "Transposes the values in a jagged array with a depth of 2. Rows become columns, columns become rows."
    Dim currentRow As Long
    Dim LowerBoundRow As Long
    Dim UpperBoundRow As Long
    Dim currentColumn As Long
    Dim LowerBoundCol As Long
    Dim UpperBoundCol As Long
    Dim Result() As Variant
    
    LowerBoundCol = LBound(SourceArray)
    UpperBoundCol = UBound(SourceArray)
    LowerBoundRow = LBound(SourceArray(LowerBoundCol))
    UpperBoundRow = UBound(SourceArray(LowerBoundCol))
    ReDim Result(LowerBoundRow To UpperBoundRow, LowerBoundCol To UpperBoundCol)
    For currentRow = LowerBoundRow To UpperBoundRow
        For currentColumn = LowerBoundCol To UpperBoundCol
            Result(currentRow, currentColumn) = SourceArray(currentColumn)(currentRow)
        Next
    Next
    Transpose_ArrayOfArrays = Result
End Function

'@Description("Returns an empty 1d array with 1 slot.")
Private Function getEmptyArray() As Variant()
Attribute getEmptyArray.VB_Description = "Returns an empty 1d array with 1 slot."
    Dim Result() As Variant
    ReDim Result(pLowerBound)
    getEmptyArray = Result
End Function

'@Description("Sorts in place a passed array using the QuickSort algorithm.")
Private Sub quickSort( _
        ByRef Arr() As Variant, _
        ByVal Low As Long, _
        ByVal High As Long, _
        ByVal ArrayType As ArrayTypes, _
        Optional ByVal Col As Long _
    )
Attribute quickSort.VB_Description = "Sorts in place a passed array using the QuickSort algorithm."
    Dim partitionIndex As Long
    If Low < High Then
        partitionIndex = qsPartition(Arr, Low, High, ArrayType, Col)
        quickSort Arr, Low, partitionIndex - 1, ArrayType, Col
        quickSort Arr, partitionIndex + 1, High, ArrayType, Col
    End If
End Sub

'@Description("Child function of quickSort. Returns the next partition index.")
Private Function qsPartition( _
        ByRef Arr() As Variant, _
        ByVal Low As Long, _
        ByVal High As Long, _
        ByVal ArrayType As ArrayTypes, _
        Optional ByVal Col As Long _
    ) As Long
Attribute qsPartition.VB_Description = "Child function of quickSort. Returns the next partition index."
    Dim i As Long
    Dim j As Long
    Dim pivot As Variant
    Dim current As Variant
    
    If ArrayType = BA_JAGGED Then
        pivot = Arr(High)(Col)
    Else
        pivot = Arr(High)
    End If
    
    i = Low - 1
    For j = Low To High - 1
        If ArrayType = BA_JAGGED Then
            current = Arr(j)(Col)
        Else
            current = Arr(j)
        End If
    
        If current <= pivot Then
            i = i + 1
            Swap Arr, i, j
        End If
    Next
    Swap Arr, i + 1, High
    qsPartition = i + 1
End Function

'@Description("Swaps the location of two elements in an array.")
Private Sub Swap(ByRef Arr() As Variant, ByVal i As Long, ByVal j As Long)
Attribute Swap.VB_Description = "Swaps the location of two elements in an array."
    Dim Element As Variant
    letOrSetElement Element, Arr(i)
    letOrSetElement Arr(i), Arr(j)
    '@Ignore UnassignedVariableUsage
    letOrSetElement Arr(j), Element
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
    Dim Result() As Variant
    If pArrayType <> BA_UNALLOCATED And pArrayType <> BA_UNDEFINED Then
        Result = pItems
        If Me.Capacity > pLength Then
            If pLength > 0 Then ReDim Preserve Result(pLowerBound To Me.UpperBound)
        End If
    Else
        Result = getEmptyArray()
    End If
    getInternalItems = Result
End Function

'@Description("Recursively parses an array to a string representation.")
Private Function recursiveToString( _
        ByRef Arr() As Variant, _
        ByVal PrettyPrint As Boolean, _
        Optional ByVal Tabs As Long = 1, _
        Optional ByVal DelimitStrings As Boolean _
    ) As String
Attribute recursiveToString.VB_Description = "Recursively parses an array to a string representation."
    Const tabWidth As Long = 2
    Dim Result As String
    Dim sep As String
    Dim i As Long

    Result = "{"
    For i = LBound(Arr) To UBound(Arr)
        sep = IIf(i = UBound(Arr), "}", IIf(PrettyPrint, ", ", ","))
        If IsArray(Arr(i)) Then
            Dim nested() As Variant
            nested = Arr(i)
            If PrettyPrint Then Result = Result & vbCrLf & Space(tabWidth * Tabs)
            Result = Result & recursiveToString(nested, PrettyPrint, Tabs + 1, DelimitStrings)
            Result = Result & IIf(i = UBound(Arr) And PrettyPrint, vbCrLf & Space(tabWidth * (Tabs - 1)) & sep, sep)
        Else
            Result = Result & Replace( _
                    CStr(getScalarRepresentation(Arr(i), DelimitStrings)), _
                    ",", _
                    vbNullString _
                ) & sep
        End If
    Next
    recursiveToString = Result
End Function

'@Description("Returns scalar values as is or appropriate representation if not scalar.")
Private Function getScalarRepresentation( _
        ByRef Element As Variant, _
        Optional ByVal DelimitStrings As Boolean _
    ) As Variant
Attribute getScalarRepresentation.VB_Description = "Returns scalar values as is or appropriate representation if not scalar."
    Dim Result As Variant
    If IsObject(Element) Then
        On Error Resume Next
        'Try to retrieve default member
        Result = Element
        On Error GoTo 0
        ' if there is no default member or it returns an object result = Empty
        If IsEmpty(Result) Then
            Result = "OBJECT"
        End If
    Else
        Result = Element
    End If
    ' if the element is array or default member returns an array then convert tostring
    If IsArray(Result) And Not IsEmpty(Result) Then
        Dim passThruArray() As Variant
        passThruArray = Result
        Result = recursiveToString(passThruArray, False, DelimitStrings:=DelimitStrings)
    End If
    If Not IsNumeric(Result) And DelimitStrings Then
        getScalarRepresentation = Chr$(34) & Result & Chr$(34)
    Else
        getScalarRepresentation = Result
    End If
    
End Function

'@Description("Returns the type of a passed array.")
Private Function getArrayType(ByVal Arr As Variant) As ArrayTypes
Attribute getArrayType.VB_Description = "Returns the type of a passed array."
    Dim Result As ArrayTypes
    If IsArray(Arr) Then
        If Not isArrayAllocated(Arr) Then
            Result = BA_UNALLOCATED
        Else
            If isMultidimensionalArray(Arr) Then
                Result = BA_MULTIDIMENSION
            ElseIf isJaggedArray(Arr) Then
                Result = BA_JAGGED
            Else
                Result = BA_ONEDIMENSION
            End If
        End If
    Else
        If IsEmpty(Arr) Then
            Result = BA_UNDEFINED
        Else
            RaiseError EC_EXPECTED_ARRAY, "getArrayType", "Arr"
        End If
    End If
    getArrayType = Result
End Function

'@Description("Tests if an array has been allocated.")
Private Function isArrayAllocated(ByVal Arr As Variant) As Boolean
Attribute isArrayAllocated.VB_Description = "Tests if an array has been allocated."
    On Error Resume Next
    isArrayAllocated = ( _
        IsArray(Arr) And _
        Not IsError(LBound(Arr, 1)) And _
        LBound(Arr, 1) <= UBound(Arr, 1) _
    )
    On Error GoTo 0
End Function

'@Description("Tests if an array is jagged.")
Private Function isJaggedArray(ByVal Arr As Variant) As Boolean
Attribute isJaggedArray.VB_Description = "Tests if an array is jagged."
    If IsArray(Arr) Then
        On Error GoTo ErrHandler
        Dim Element As Variant
        For Each Element In Arr
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
Private Function isMultidimensionalArray(ByVal Arr As Variant) As Boolean
Attribute isMultidimensionalArray.VB_Description = "Tests if an array is multidimensioned."
    If IsArray(Arr) Then
        On Error GoTo ErrHandler
        '@Ignore VariableNotUsed
        Dim localUpperBound As Long
        '@Ignore AssignmentNotUsed
        localUpperBound = UBound(Arr, 2)
        isMultidimensionalArray = True
        On Error GoTo 0
    End If
    Exit Function
ErrHandler:
    Err.Clear
End Function

'@Ignore FunctionReturnValueNotUsed
'@Description("re-indexes array to conform to new bounds. Does not change length.")
Private Function Rebase( _
        Optional ByRef Recurse As Boolean, _
        Optional ByVal Arr As Variant _
    ) As Variant()
Attribute Rebase.VB_Description = "re-indexes array to conform to new bounds. Does not change length."
    Dim isMD As Boolean
    Dim applyToMe As Boolean
    Dim newItems() As Variant
    
    If IsMissing(Arr) Or Not IsArray(Arr) Then applyToMe = True
    newItems = IIf(applyToMe, pItems, Arr)
    isMD = isMultidimensionalArray(newItems)
    If isMD Then
        newItems = multiToJagged(newItems)
        Recurse = True
    End If
    newItems = recursiveRebase(newItems, Recurse)
    If isMD Then newItems = jaggedToMulti(newItems)
    If applyToMe Then pItems = newItems
    Rebase = newItems
End Function

'@Description("re-indexes nested arrays to conform to new bounds. Child function of Rebase)
Private Function recursiveRebase( _
        ByRef Arr() As Variant, _
        ByVal Recurse As Boolean _
    ) As Variant()
    Dim i As Long
    Dim localLowerBound As Long
    Dim localUpperBound As Long
    Dim offset As Long
    Dim newItems() As Variant
    
    localLowerBound = LBound(Arr)
    localUpperBound = UBound(Arr)
    offset = pLowerBound - localLowerBound
    ReDim newItems(pLowerBound To localUpperBound + offset)
    For i = localLowerBound To localUpperBound
        If IsArray(Arr(i)) And Recurse Then
            Dim nested() As Variant
            nested = Arr(i)
            newItems(i + offset) = recursiveRebase(nested, Recurse)
        Else
            newItems(i + offset) = Arr(i)
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
Private Function jaggedToMulti( _
        ByRef Arr() As Variant, _
        Optional ByVal Depth As Long, _
        Optional ByVal CurrentDepth As Long, _
        Optional ByRef Crumbs As Variant, _
        Optional ByRef Result As Variant, _
        Optional ByVal EnsureScalar As Boolean _
    ) As Variant()
Attribute jaggedToMulti.VB_Description = "Converts jagged arrays to multidimension arrays"
    Dim i As Long
    Dim localLowerBound As Long
    Dim localUpperBound As Long
    Dim localDepth As Long
    Dim localCurrentDepth As Long
    Dim localCrumbs() As Variant
    Dim localResult() As Variant
    
    
    localDepth = Depth
    localCurrentDepth = CurrentDepth
    
    If localDepth = 0 Then
        If Not isJaggedArray(Arr) Then
            RaiseError EC_EXPECTED_JAGGED_ARRAY, "jaggedToMulti", "Arr()"
            Exit Function
        End If
        localDepth = getJaggedArrayDepth(Arr)
    End If
    
    If localDepth > 20 Then
        ' too many dimensions to be handled by getElementByBreadcrumb
        RaiseError EC_MAX_DIMENSIONS_LIMIT, "jaggedToMulti()", "arr()"
    ElseIf localDepth > 1 Then
    
        If IsMissing(Crumbs) Then
            ReDim localCrumbs(localDepth - 1)
        Else
            localCrumbs = Crumbs
        End If
        
        If IsMissing(Result) Then
            localResult = createMultidimensionalArray(mapJaggedArray(Arr, KnownDepth:=localDepth))
        Else
            localResult = Result
        End If
        
        localCurrentDepth = localCurrentDepth + 1
        localLowerBound = LBound(Arr)
        localUpperBound = UBound(Arr)
        For i = localLowerBound To localUpperBound
            localCrumbs(localCurrentDepth - 1) = i
            If localCurrentDepth = localDepth Or (Not IsArray(Arr(i)) And Not EnsureScalar) Then
                If EnsureScalar Then
                    localResult = letElementByBreadcrumb(localResult, localCrumbs, getScalarRepresentation(Arr(i)))
                Else
                    localResult = letElementByBreadcrumb(localResult, localCrumbs, Arr(i))
                End If
            Else
                Dim nested() As Variant
                nested = Arr(i)
                localResult = jaggedToMulti( _
                    Arr:=nested, _
                    Depth:=localDepth, _
                    CurrentDepth:=localCurrentDepth, _
                    Crumbs:=localCrumbs, _
                    Result:=localResult, _
                    EnsureScalar:=EnsureScalar _
                )
            End If
        Next
    Else
        localResult = Arr
    End If
    
    jaggedToMulti = localResult
End Function

'@Description("Converts multidimension arrays to jagged arrays.")
Private Function multiToJagged( _
        ByRef Arr() As Variant, _
        Optional ByVal Depth As Long, _
        Optional ByVal CurrentDepth As Long, _
        Optional ByRef Crumbs As Variant _
    ) As Variant()
Attribute multiToJagged.VB_Description = "Converts multidimension arrays to jagged arrays."
    Dim i As Long
    Dim localLowerBound As Long
    Dim localUpperBound As Long
    Dim localDepth As Long
    Dim localCurrentDepth As Long
    Dim localCrumbs() As Variant
    Dim Result() As Variant
    
    localDepth = Depth
    localCurrentDepth = CurrentDepth
    
    If localDepth = 0 Then
        If Not isMultidimensionalArray(Arr) Then
            RaiseError EC_EXPECTED_MULTIDIMENSION_ARRAY, "multiToJagged", "Arr()"
            Exit Function
        End If
        localDepth = getMultidimensionalArrayDepth(Arr)
    End If
    
    If localDepth > 20 Then
        RaiseError EC_MAX_DIMENSIONS_LIMIT, "multiToJagged", "Arr()"
    ElseIf localDepth > 1 Then
        If IsMissing(Crumbs) Then
            ReDim localCrumbs(localDepth - 1)
        Else
            localCrumbs = Crumbs
        End If
        localCurrentDepth = localCurrentDepth + 1
        localLowerBound = LBound(Arr, localCurrentDepth)
        localUpperBound = UBound(Arr, localCurrentDepth)
        ReDim Result(localLowerBound To localUpperBound)
        For i = localLowerBound To localUpperBound
            localCrumbs(localCurrentDepth - 1) = i
            If localCurrentDepth = localDepth Then
                Result(i) = getElementByBreadcrumb(Arr, localCrumbs)
            Else
                Result(i) = multiToJagged(Arr, localDepth, localCurrentDepth, localCrumbs)
            End If
        Next
    Else
        Result = Arr
    End If
    multiToJagged = Result
End Function

'@Description("Returns a number representing how many levels of nested arrays are contained in array.")
Private Function getJaggedArrayDepth(ByRef Arr() As Variant) As Long
Attribute getJaggedArrayDepth.VB_Description = "Returns a number representing how many levels of nested arrays are contained in array."
    Dim i As Long
    Dim localLowerBound As Long
    Dim localUpperBound As Long
    Dim CurrentDepth As Long
    Dim maxDepth As Long
    Dim Depth As Long
    
    If IsArray(Arr) Then
        Depth = Depth + 1
        localLowerBound = LBound(Arr)
        localUpperBound = UBound(Arr)
        For i = localLowerBound To localUpperBound
            If IsArray(Arr(i)) Then
                Dim nested() As Variant
                nested = Arr(i)
                CurrentDepth = getJaggedArrayDepth(nested)
            End If
            If CurrentDepth > maxDepth Then maxDepth = CurrentDepth
        Next
        Depth = Depth + maxDepth
    End If
    getJaggedArrayDepth = Depth
End Function

'@Description("Gets the maximium bound values for a specified dimension of a jagged array.")
Private Function getMaxBoundsAtDimension( _
        ByRef SourceArray() As Variant, _
        Optional ByVal Dimension As Long = 1, _
        Optional ByVal currentDimension As Long = 1 _
    ) As Variant()
Attribute getMaxBoundsAtDimension.VB_Description = "Gets the maximium bound values for a specified dimension of a jagged array."
    Dim maxBounds() As Variant
    
    If currentDimension < Dimension Then
        Dim i As Long
        Dim currentBounds() As Variant
        Dim localLowerBound As Long
        Dim localUpperBound As Long
        
        localLowerBound = LBound(SourceArray)
        localUpperBound = UBound(SourceArray)
        For i = localLowerBound To localUpperBound
            If IsArray(SourceArray(i)) Then
                Dim nested() As Variant
                nested = SourceArray(i)
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
        maxBounds = getArrayBounds(SourceArray)
    End If
    
    getMaxBoundsAtDimension = maxBounds
End Function

'@Description("Returns an array representing the structure of a jagged array")
Private Function mapJaggedArray( _
        ByRef SourceArray() As Variant, _
        Optional ByVal KnownDepth As Long _
    ) As Variant()
Attribute mapJaggedArray.VB_Description = "Returns an array representing the structure of a jagged array"
    Dim i As Long
    Dim localMap() As Variant
    Dim localDepth As Long

    If Not IsArray(SourceArray) Then Exit Function

    If KnownDepth > 0 Then
        localDepth = KnownDepth
    Else
        localDepth = getJaggedArrayDepth(SourceArray)
    End If

    ' crumbs are base 0
    ReDim localMap(0 To localDepth - 1)

    For i = LBound(localMap) To UBound(localMap)
        localMap(i) = getMaxBoundsAtDimension(SourceArray, i + 1)
    Next

    mapJaggedArray = localMap
End Function

'@Description("Returns a 0-based array with two slots. First slot is the lowerbound of the passed array. second slot (index 0) is the upper bound.")
Private Function getArrayBounds(ByVal Arr As Variant) As Variant()
Attribute getArrayBounds.VB_Description = "Returns a 0-based array with two slots. First slot is the lowerbound of the passed array. second slot (index 0) is the upper bound."
    Dim Result(0 To 1) As Variant
    If IsArray(Arr) Then
        Result(0) = LBound(Arr)
        Result(1) = UBound(Arr)
    End If
    getArrayBounds = Result
End Function

'@Description("Returns a number representing the total number of dimensions of a multidimension array.")
Private Function getMultidimensionalArrayDepth(ByVal Arr As Variant) As Long
Attribute getMultidimensionalArrayDepth.VB_Description = "Returns a number representing the total number of dimensions of a multidimension array."
    Dim i As Long
    '@Ignore VariableNotUsed
    Dim void As Long
    
    On Error Resume Next
    Do
        i = i + 1
        void = UBound(Arr, i)
    Loop Until Err.number <> 0
    Err.Clear
    On Error GoTo 0
    getMultidimensionalArrayDepth = i - 1
End Function

'@Description("Returns the element stored in an array based on a crumb argument. A crumb is an array representing the location of the multidimension array slot.")
Private Function getElementByBreadcrumb( _
        ByRef Arr() As Variant, _
        ByRef Crumb() As Variant _
    ) As Variant
Attribute getElementByBreadcrumb.VB_Description = "Returns the element stored in an array based on a crumb argument. A crumb is an array representing the location of the multidimension array slot."
    Dim Result As Variant
    Select Case UBound(Crumb)
    Case 0
        letOrSetElement Result, Arr(Crumb(0))
    Case 1
        letOrSetElement Result, Arr(Crumb(0), Crumb(1))
    Case 2
        letOrSetElement Result, Arr(Crumb(0), Crumb(1), Crumb(2))
    Case 3
        letOrSetElement Result, Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3))
    Case 4
        letOrSetElement Result, Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4))
    Case 5
        letOrSetElement Result, Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5))
    Case 6
        letOrSetElement Result, Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6))
    Case 7
        letOrSetElement Result, Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7))
    Case 8
        letOrSetElement Result, Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8))
    Case 9
        letOrSetElement Result, Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), _
            Crumb(9))
    Case 10
        letOrSetElement Result, Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), _
            Crumb(9), Crumb(10))
    Case 11
        letOrSetElement Result, Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), _
            Crumb(9), Crumb(10), Crumb(11))
    Case 12
        letOrSetElement Result, Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), _
            Crumb(9), Crumb(10), Crumb(11), Crumb(12))
    Case 13
        letOrSetElement Result, Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), _
            Crumb(9), Crumb(10), Crumb(11), Crumb(12), Crumb(13))
    Case 14
        letOrSetElement Result, Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), _
            Crumb(9), Crumb(10), Crumb(11), Crumb(12), Crumb(13), Crumb(14))
    Case 15
        letOrSetElement Result, Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), _
            Crumb(9), Crumb(10), Crumb(11), Crumb(12), Crumb(13), Crumb(14), Crumb(15))
    Case 16
        letOrSetElement Result, Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), _
            Crumb(9), Crumb(10), Crumb(11), Crumb(12), Crumb(13), Crumb(14), Crumb(15), Crumb(16))
    Case 17
        letOrSetElement Result, Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), _
            Crumb(9), Crumb(10), Crumb(11), Crumb(12), Crumb(13), Crumb(14), Crumb(15), Crumb(16), Crumb(17))
    Case 18
        letOrSetElement Result, Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), _
            Crumb(9), Crumb(10), Crumb(11), Crumb(12), Crumb(13), Crumb(14), Crumb(15), Crumb(16), Crumb(17), Crumb(18))
    Case 19
        letOrSetElement Result, Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), _
            Crumb(9), Crumb(10), Crumb(11), Crumb(12), Crumb(13), Crumb(14), Crumb(15), Crumb(16), Crumb(17), Crumb(18), Crumb(19))
    Case 20
        letOrSetElement Result, Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), _
            Crumb(9), Crumb(10), Crumb(11), Crumb(12), Crumb(13), Crumb(14), Crumb(15), Crumb(16), Crumb(17), Crumb(18), Crumb(19), Crumb(20))
    End Select
    '@Ignore UnassignedVariableUsage
    getElementByBreadcrumb = Result
End Function

'@Description("Assigns an element to a location in a multidimensonal array based on a crumb argument. A crumb is an array representing the location of the multidimension array slot.")
Private Function letElementByBreadcrumb( _
        ByRef Arr() As Variant, _
        ByRef Crumb() As Variant, _
        ByVal Element As Variant _
    ) As Variant
Attribute letElementByBreadcrumb.VB_Description = "Assigns an element to a location in a multidimensonal array based on a crumb argument. A crumb is an array representing the location of the multidimension array slot."
    Select Case UBound(Crumb)
    Case 0
        letOrSetElement Arr(Crumb(0)), Element
    Case 1
        letOrSetElement Arr(Crumb(0), Crumb(1)), Element
    Case 2
        letOrSetElement Arr(Crumb(0), Crumb(1), Crumb(2)), Element
    Case 3
        letOrSetElement Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3)), Element
    Case 4
        letOrSetElement Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4)), Element
    Case 5
        letOrSetElement Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5)), Element
    Case 6
        letOrSetElement Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6)), Element
    Case 7
        letOrSetElement Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7)), Element
    Case 8
        letOrSetElement Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8)), Element
    Case 9
        letOrSetElement Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), Crumb(9)), _
            Element
    Case 10
        letOrSetElement Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), Crumb(9), _
            Crumb(10)), Element
    Case 11
        letOrSetElement Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), Crumb(9), _
            Crumb(10), Crumb(11)), Element
    Case 12
        letOrSetElement Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), Crumb(9), _
            Crumb(10), Crumb(11), Crumb(12)), Element
    Case 13
        letOrSetElement Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), Crumb(9), _
            Crumb(10), Crumb(11), Crumb(12), Crumb(13)), Element
    Case 14
        letOrSetElement Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), Crumb(9), _
            Crumb(10), Crumb(11), Crumb(12), Crumb(13), Crumb(14)), Element
    Case 15
        letOrSetElement Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), Crumb(9), _
            Crumb(10), Crumb(11), Crumb(12), Crumb(13), Crumb(14), Crumb(15)), Element
    Case 16
        letOrSetElement Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), Crumb(9), _
        Crumb(10), Crumb(11), Crumb(12), Crumb(13), Crumb(14), Crumb(15), Crumb(16)), Element
    Case 17
        letOrSetElement Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), Crumb(9), _
            Crumb(10), Crumb(11), Crumb(12), Crumb(13), Crumb(14), Crumb(15), Crumb(16), Crumb(17)), Element
    Case 18
        letOrSetElement Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), Crumb(9), _
            Crumb(10), Crumb(11), Crumb(12), Crumb(13), Crumb(14), Crumb(15), Crumb(16), Crumb(17), Crumb(18)), Element
    Case 19
        letOrSetElement Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), Crumb(9), _
        Crumb(10), Crumb(11), Crumb(12), Crumb(13), Crumb(14), Crumb(15), Crumb(16), Crumb(17), Crumb(18), Crumb(19)), Element
    Case 20
        letOrSetElement Arr(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), Crumb(9), _
        Crumb(10), Crumb(11), Crumb(12), Crumb(13), Crumb(14), Crumb(15), Crumb(16), Crumb(17), Crumb(18), Crumb(19), Crumb(20)), _
        Element
    End Select

    letElementByBreadcrumb = Arr
End Function

'@Description("Returns an empty multidimensional array with the dimensions specified in the crumb argument. A crumb is an array representing the structure of the multidimension array.")
Private Function createMultidimensionalArray(ByRef Crumb() As Variant) As Variant()
Attribute createMultidimensionalArray.VB_Description = "Returns an empty multidimensional array with the dimensions specified in the crumb argument. A crumb is an array representing the structure of the multidimension array."
    Dim Result() As Variant
    Select Case UBound(Crumb)
    Case 0
        ReDim Result(Crumb(0)(0) To Crumb(0)(1))
    Case 1
        ReDim Result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1))
    Case 2
        ReDim Result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1))
    Case 3
        ReDim Result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1))
    Case 4
        ReDim Result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1))
    Case 5
        ReDim Result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1), Crumb(5)(0) To Crumb(5)(1))
    Case 6
        ReDim Result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1), Crumb(5)(0) To Crumb(5)(1), Crumb(6)(0) To Crumb(6)(1))
    Case 7
        ReDim Result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1), Crumb(5)(0) To Crumb(5)(1), Crumb(6)(0) To Crumb(6)(1), Crumb(7)(0) To Crumb(7)(1))
    Case 8
        ReDim Result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1), Crumb(5)(0) To Crumb(5)(1), Crumb(6)(0) To Crumb(6)(1), Crumb(7)(0) To Crumb(7)(1), _
                     Crumb(8)(0) To Crumb(8)(1))
    Case 9
        ReDim Result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1), Crumb(5)(0) To Crumb(5)(1), Crumb(6)(0) To Crumb(6)(1), Crumb(7)(0) To Crumb(7)(1), _
                     Crumb(8)(0) To Crumb(8)(1), Crumb(9)(0) To Crumb(9)(1))
    Case 10
        ReDim Result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1), Crumb(5)(0) To Crumb(5)(1), Crumb(6)(0) To Crumb(6)(1), Crumb(7)(0) To Crumb(7)(1), _
                     Crumb(8)(0) To Crumb(8)(1), Crumb(9)(0) To Crumb(9)(1), Crumb(10)(0) To Crumb(10)(1))
    Case 11
        ReDim Result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1), Crumb(5)(0) To Crumb(5)(1), Crumb(6)(0) To Crumb(6)(1), Crumb(7)(0) To Crumb(7)(1), _
                     Crumb(8)(0) To Crumb(8)(1), Crumb(9)(0) To Crumb(9)(1), Crumb(10)(0) To Crumb(10)(1), Crumb(11)(0) To Crumb(11)(1))
    Case 12
        ReDim Result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1), Crumb(5)(0) To Crumb(5)(1), Crumb(6)(0) To Crumb(6)(1), Crumb(7)(0) To Crumb(7)(1), _
                     Crumb(8)(0) To Crumb(8)(1), Crumb(9)(0) To Crumb(9)(1), Crumb(10)(0) To Crumb(10)(1), Crumb(11)(0) To Crumb(11)(1), _
                     Crumb(12)(0) To Crumb(12)(1))
    Case 13
        ReDim Result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1), Crumb(5)(0) To Crumb(5)(1), Crumb(6)(0) To Crumb(6)(1), Crumb(7)(0) To Crumb(7)(1), _
                     Crumb(8)(0) To Crumb(8)(1), Crumb(9)(0) To Crumb(9)(1), Crumb(10)(0) To Crumb(10)(1), Crumb(11)(0) To Crumb(11)(1), _
                     Crumb(12)(0) To Crumb(12)(1), Crumb(13)(0) To Crumb(13)(1))
    Case 14
        ReDim Result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1), Crumb(5)(0) To Crumb(5)(1), Crumb(6)(0) To Crumb(6)(1), Crumb(7)(0) To Crumb(7)(1), _
                     Crumb(8)(0) To Crumb(8)(1), Crumb(9)(0) To Crumb(9)(1), Crumb(10)(0) To Crumb(10)(1), Crumb(11)(0) To Crumb(11)(1), _
                     Crumb(12)(0) To Crumb(12)(1), Crumb(13)(0) To Crumb(13)(1), Crumb(14)(0) To Crumb(14)(1))
    Case 15
        ReDim Result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1), Crumb(5)(0) To Crumb(5)(1), Crumb(6)(0) To Crumb(6)(1), Crumb(7)(0) To Crumb(7)(1), _
                     Crumb(8)(0) To Crumb(8)(1), Crumb(9)(0) To Crumb(9)(1), Crumb(10)(0) To Crumb(10)(1), Crumb(11)(0) To Crumb(11)(1), _
                     Crumb(12)(0) To Crumb(12)(1), Crumb(13)(0) To Crumb(13)(1), Crumb(14)(0) To Crumb(14)(1), Crumb(15)(0) To Crumb(15)(1))
    Case 16
        ReDim Result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1), Crumb(5)(0) To Crumb(5)(1), Crumb(6)(0) To Crumb(6)(1), Crumb(7)(0) To Crumb(7)(1), _
                     Crumb(8)(0) To Crumb(8)(1), Crumb(9)(0) To Crumb(9)(1), Crumb(10)(0) To Crumb(10)(1), Crumb(11)(0) To Crumb(11)(1), _
                     Crumb(12)(0) To Crumb(12)(1), Crumb(13)(0) To Crumb(13)(1), Crumb(14)(0) To Crumb(14)(1), Crumb(15)(0) To Crumb(15)(1), _
                     Crumb(16)(0) To Crumb(16)(1))
    Case 17
        ReDim Result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1), Crumb(5)(0) To Crumb(5)(1), Crumb(6)(0) To Crumb(6)(1), Crumb(7)(0) To Crumb(7)(1), _
                     Crumb(8)(0) To Crumb(8)(1), Crumb(9)(0) To Crumb(9)(1), Crumb(10)(0) To Crumb(10)(1), Crumb(11)(0) To Crumb(11)(1), _
                     Crumb(12)(0) To Crumb(12)(1), Crumb(13)(0) To Crumb(13)(1), Crumb(14)(0) To Crumb(14)(1), Crumb(15)(0) To Crumb(15)(1), _
                     Crumb(16)(0) To Crumb(16)(1), Crumb(17)(0) To Crumb(17)(1))
    Case 18
        ReDim Result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1), Crumb(5)(0) To Crumb(5)(1), Crumb(6)(0) To Crumb(6)(1), Crumb(7)(0) To Crumb(7)(1), _
                     Crumb(8)(0) To Crumb(8)(1), Crumb(9)(0) To Crumb(9)(1), Crumb(10)(0) To Crumb(10)(1), Crumb(11)(0) To Crumb(11)(1), _
                     Crumb(12)(0) To Crumb(12)(1), Crumb(13)(0) To Crumb(13)(1), Crumb(14)(0) To Crumb(14)(1), Crumb(15)(0) To Crumb(15)(1), _
                     Crumb(16)(0) To Crumb(16)(1), Crumb(17)(0) To Crumb(17)(1), Crumb(18)(0) To Crumb(18)(1))
    Case 19
        ReDim Result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1), Crumb(5)(0) To Crumb(5)(1), Crumb(6)(0) To Crumb(6)(1), Crumb(7)(0) To Crumb(7)(1), _
                     Crumb(8)(0) To Crumb(8)(1), Crumb(9)(0) To Crumb(9)(1), Crumb(10)(0) To Crumb(10)(1), Crumb(11)(0) To Crumb(11)(1), _
                     Crumb(12)(0) To Crumb(12)(1), Crumb(13)(0) To Crumb(13)(1), Crumb(14)(0) To Crumb(14)(1), Crumb(15)(0) To Crumb(15)(1), _
                     Crumb(16)(0) To Crumb(16)(1), Crumb(17)(0) To Crumb(17)(1), Crumb(18)(0) To Crumb(18)(1), Crumb(19)(0) To Crumb(19)(1))
    Case 20
        ReDim Result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1), Crumb(5)(0) To Crumb(5)(1), Crumb(6)(0) To Crumb(6)(1), Crumb(7)(0) To Crumb(7)(1), _
                     Crumb(8)(0) To Crumb(8)(1), Crumb(9)(0) To Crumb(9)(1), Crumb(10)(0) To Crumb(10)(1), Crumb(11)(0) To Crumb(11)(1), _
                     Crumb(12)(0) To Crumb(12)(1), Crumb(13)(0) To Crumb(13)(1), Crumb(14)(0) To Crumb(14)(1), Crumb(15)(0) To Crumb(15)(1), _
                     Crumb(16)(0) To Crumb(16)(1), Crumb(17)(0) To Crumb(17)(1), Crumb(18)(0) To Crumb(18)(1), Crumb(19)(0) To Crumb(19)(1), _
                     Crumb(20)(0) To Crumb(20)(1))
    End Select
    createMultidimensionalArray = Result
End Function






















