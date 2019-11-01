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
Private Const OBJECT_REPR As String = "OBJECT"
Private Const MISSING_LONG As Long = -9999

Private Enum ArrayTypes
    BA_UNDEFINED
    BA_UNALLOCATED
    BA_ONEDIMENSION
    BA_MULTIDIMENSION
    BA_JAGGED
End Enum

' Public scope as used in unit tests
Public Enum ErrorCodes
    EC_START = vbObjectError + 512
    EC_EXPECTED_RANGE_OBJECT
    EC_EXPECTED_COLLECTION_OBJECT
    EC_MAX_DIMENSIONS_LIMIT
    EC_EXCEEDS_MAX_SORT_DEPTH
    EC_EXPECTED_JAGGED_ARRAY
    EC_EXPECTED_MULTIDIMENSION_ARRAY
    EC_EXPECTED_ARRAY
    EC_NULL_STRING
    EC_UNALLOCATED_ARRAY
    EC_UNDEFINED_ARRAY
    EC_INVALID_MULTIDIMENSIONAL_ARRAY_OPERATION
    EC_EXPECTED_VARIANT_ARRAY
    EC_END
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
Private ErrorDefinitions(EC_START To EC_END) As ErrorDefinition
Private LowerBoundSet As Boolean

'@Description("Constructor")
Private Sub Class_Initialize()
Attribute Class_Initialize.VB_Description = "Constructor"
    Me.Capacity = DEFAULT_CAPACITY
    pArrayType = BA_UNALLOCATED
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
Public Property Let Capacity(ByVal value As Long)
Attribute Capacity.VB_Description = "Sets the capacity of the internal array."
    If value < 0 Then Err.Raise 9
    If value <> pCapacity Then
        If pCapacity > 0 Then
            If GetArrayLength(pItems) <> value Then
                Dim newItems() As Variant
                newItems = pItems
                ReDim Preserve newItems(pLowerBound To (value + pLowerBound - 1))
                pItems = newItems
            End If
        Else
            ReDim pItems(pLowerBound To (DEFAULT_CAPACITY + pLowerBound - 1))
        End If
        pCapacity = UBound(pItems) - pLowerBound + 1
    End If
End Property

'@Description("Returns the number of entries stored in the array.")
Public Property Get length() As Long
Attribute length.VB_Description = "Returns the number of entries stored in the array."
    length = pLength
End Property

'@Description("Returns largest used index in the array.")
Public Property Get UpperBound()
Attribute UpperBound.VB_Description = "Returns largest used index in the array."
    If pArrayType = BA_UNALLOCATED Then
        UpperBound = -1
    Else
        UpperBound = pLength + pLowerBound - 1
    End If
End Property

'@Description("Returns smallest used index in the array.")
Public Property Get LowerBound() As Long
Attribute LowerBound.VB_Description = "Returns smallest used index in the array."
    LowerBound = pLowerBound
End Property

'@Description("Sets the starting index of the array. If an array is already stored then it will be re-indexed to match the new value")
Public Property Let LowerBound(ByVal value As Long)
Attribute LowerBound.VB_Description = "Sets the starting index of the array. If an array is already stored then it will be re-indexed to match the new value"
    LowerBoundSet = True
    If value <> pLowerBound And value >= 0 Then
        pLowerBound = value
        pItems = Rebase()
        pCapacity = pLength
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
Public Property Let Item(ByVal index As Long, ByVal element As Variant)
Attribute Item.VB_Description = "Assigns the passed element to the array at the specified index."
    If Me.UpperBound >= index Then
        If index < Me.LowerBound Then
            Me.Unshift element
        Else
            LetOrSetElement pItems(index), element
        End If
    Else
        Me.Push element
    End If
End Property

'@Description("Returns the stored array.")
Public Property Get Items() As Variant
Attribute Items.VB_Description = "Returns the stored array."
    Dim result() As Variant
    result = GetInternalItems()
    If pArrayType = ArrayTypes.BA_MULTIDIMENSION Then
        If IsJaggedArray(result) Then result = JaggedToMulti(result)
    End If
    Items = result
End Property

'@Description("Stores the passed array.")
Public Property Let Items(ByVal values As Variant)
Attribute Items.VB_Description = "Stores the passed array."
    Dim localLowerBound As Long
    Dim localValues() As Variant
    If IsArray(values) Then
        If TypeName(values) <> "Variant()" Then
            RaiseError EC_EXPECTED_VARIANT_ARRAY, "Items", "Values"
            Exit Property
        End If
        localValues = values
        pArrayType = GetArrayType(localValues)
        
        If pArrayType = ArrayTypes.BA_UNDEFINED Then
            RaiseError EC_EXPECTED_ARRAY, "Items", "Values"
        Else
            If pArrayType = BA_UNALLOCATED Then
                localValues = GetEmptyArray()
            End If
            
            localLowerBound = LBound(localValues)
            
            If Not LowerBoundSet Then
                pLowerBound = localLowerBound
            End If
            
            If pArrayType = BA_MULTIDIMENSION Then
                localValues = MultiToJagged(localValues)
            End If
            
            If localLowerBound <> pLowerBound Then
                pItems = Rebase(localValues, pArrayType)
            Else
                pItems = localValues
            End If
            pLength = GetArrayLength(localValues)
            Me.Capacity = pLength
        End If
    Else
        ' If Values is not an array then clear our internal array and push the Values arg
        ' TODO: Document and write tests
        Me.Clear.Push values
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
    Dim element As Variant

    If pArrayType = ArrayTypes.BA_UNALLOCATED Or _
       pArrayType = ArrayTypes.BA_UNDEFINED Then
        pArrayType = ArrayTypes.BA_ONEDIMENSION
    End If

    For Each element In args
        If pLength = pCapacity Then
            EnsureCapacity pLength + 1
        End If
        If IsArray(element) Then
            Dim arrayElement() As Variant
            Dim ArrayElementType As ArrayTypes
            arrayElement = element
            If pArrayType = BA_ONEDIMENSION Then
                ArrayElementType = GetArrayType(arrayElement)
                If ArrayElementType = BA_MULTIDIMENSION Then
                    pArrayType = BA_MULTIDIMENSION
                Else
                    pArrayType = BA_JAGGED
                End If
            End If
            If LBound(arrayElement) <> pLowerBound Then
                arrayElement = Rebase(arrayElement, ArrayElementType)
            End If
            LetOrSetElement pItems(pLength + pLowerBound), arrayElement
        Else
            LetOrSetElement pItems(pLength + pLowerBound), element
'            If pArrayType = BA_MULTIDIMENSION Then
'                'encase in array if pushing to md
'                letOrSetElement pItems(pLength + pLowerBound), Array(element)
'            Else
'                letOrSetElement pItems(pLength + pLowerBound), element
'            End If
        End If
        pLength = pLength + 1
    Next
    Push = pLength
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Removes the last element from the array and returns that element.")
Public Function Pop() As Variant
Attribute Pop.VB_Description = "Removes the last element from the array and returns that element."
    Dim result As Variant
    Dim newItems() As Variant
    If pLength > 0 Then
        result = pItems(Me.UpperBound)
        newItems = Me.Slice(pLowerBound, Me.UpperBound)
        Me.Items = newItems
    End If
    Pop = result
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Removes the first element from the array and returns that element.")
Public Function Shift() As Variant
Attribute Shift.VB_Description = "Removes the first element from the array and returns that element."
    Dim newItems() As Variant
    Dim result As Variant
    If pLength > 0 Then
        result = pItems(pLowerBound)
        newItems = Me.Slice(pLowerBound + 1)
        Me.Items = newItems
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
    oldItems = GetInternalItems()
    Me.Items = newItems
    If oldType <> BA_UNALLOCATED And oldType <> BA_UNDEFINED Then
        Me.Concat oldItems
        pArrayType = oldType
    End If
    Unshift = Me.length
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Returns a string representing the array structure and its elements.")
Public Function ToString( _
        Optional ByVal prettyPrint As Boolean, _
        Optional ByVal delimitStrings As Boolean, _
        Optional ByVal openingDelimiter As String, _
        Optional ByVal closingDelimiter As String _
    ) As String
Attribute ToString.VB_Description = "Returns a string representing the array structure and its elements."
    Dim arrayType As ArrayTypes
    Dim result As String
    Dim localItems() As Variant
    
    localItems = GetInternalItems()
    arrayType = GetArrayType(localItems)
    If Not arrayType = ArrayTypes.BA_UNDEFINED And _
       Not arrayType = ArrayTypes.BA_UNALLOCATED Then
        If arrayType = ArrayTypes.BA_MULTIDIMENSION Then
            localItems = MultiToJagged(localItems)
        End If
        result = RecursiveToString( _
            arr:=localItems, _
            prettyPrint:=prettyPrint, _
            delimitStrings:=delimitStrings, _
            openingDelimiter:=openingDelimiter, _
            closingDelimiter:=closingDelimiter _
        )
    End If
    ToString = result
End Function

' TODO: add param for recursion
'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Determines whether the SearchArrayay includes a certain value among its entries.")
Public Function Includes( _
        ByVal searchElement As Variant, _
        Optional ByVal fromIndex As Long = MISSING_LONG _
    ) As Boolean
Attribute Includes.VB_Description = "Determines whether the SearchArrayay includes a certain value among its entries."
    Dim localLength As Long
    Dim currentIndex As Long
    Dim searchArray() As Variant
    
    searchArray = GetInternalItems()
    localLength = Me.length
    
    If localLength = 0 Then
        Includes = False
        Exit Function
    End If
        
    If fromIndex > Me.LowerBound Then
        currentIndex = fromIndex
    ElseIf fromIndex = MISSING_LONG Then
        currentIndex = Me.LowerBound
    Else
        currentIndex = localLength + fromIndex
    End If
    
    Includes = RecursiveIncludes(searchElement, searchArray, currentIndex)
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Returns a new array that contains the keys for each index in the array.")
Public Function Keys() As Variant()
Attribute Keys.VB_Description = "Returns a new array that contains the keys for each index in the array."
    Dim i As Long
    Dim localLowerBound As Long
    Dim result() As Variant
    
    If pArrayType = BA_UNDEFINED Then
        RaiseError EC_UNDEFINED_ARRAY, "Keys"
    ElseIf pArrayType = BA_UNALLOCATED Then
        RaiseError EC_UNALLOCATED_ARRAY, "Keys"
    Else
        localLowerBound = Me.LowerBound
        ReDim result(0 To Me.length - 1)
        For i = LBound(result) To UBound(result)
            result(i) = i + localLowerBound
        Next
    End If
    Keys = result
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Returns the largest value in a list of values. If arguments are provided it will return the largest argument. If no arguments are provided it will return the largest element in the stored array")
Public Function Max(ParamArray args() As Variant) As Variant
Attribute Max.VB_Description = "Returns the largest value in a list of values. If arguments are provided it will return the largest argument. If no arguments are provided it will return the largest element in the stored array"
    Dim arr() As Variant
    
    If UBound(args) < LBound(args) Then          'Test if no args provided
        arr = GetInternalItems()
    Else
        arr = args
    End If
    
    Max = RecursiveMax(arr)
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Returns the smallest value in a list of values. If arguments are provided it will return the smallest argument. If no arguments are provided it will return the smallest element in the stored array")
Public Function Min(ParamArray args() As Variant) As Variant
Attribute Min.VB_Description = "Returns the smallest value in a list of values. If arguments are provided it will return the smallest argument. If no arguments are provided it will return the smallest element in the stored array"
    Dim arr() As Variant
    
    If UBound(args) < LBound(args) Then          'Test if no args provided
        arr = GetInternalItems()
    Else
        arr = args
    End If
    
    Min = RecursiveMin(arr)
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Returns a new array containing the portion of the stored array between the StartIndex and EndIndex")
Public Function Slice( _
        ByVal startIndex As Long, _
        Optional ByVal endIndex As Long = MISSING_LONG _
    ) As Variant()
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
    
    arr = GetInternalItems()
    localLength = Me.length
    relativeStart = startIndex
    If relativeStart < LBound(arr) Then
        If relativeStart < 0 Then
            oldIndex = Max((localLength + relativeStart), LBound(arr))
        Else
            oldIndex = Max((localLength - relativeStart), LBound(arr))
        End If
    Else
        oldIndex = Min(relativeStart, localLength)
    End If
    If endIndex = MISSING_LONG Then
        relativeEnd = localLength + pLowerBound
    Else
        relativeEnd = endIndex
    End If
    If relativeEnd < LBound(arr) Then
        final = Max((localLength + relativeEnd), LBound(arr))
    Else
        final = Min(relativeEnd, localLength + pLowerBound)
    End If
    
    newIndex = pLowerBound
    count = Max(final - oldIndex, 0) + pLowerBound
    If count > newIndex Then
        ReDim result(newIndex To count - 1)
        Do While oldIndex < final
            If oldIndex >= LBound(arr) And oldIndex <= UBound(arr) Then
                LetOrSetElement result(newIndex), arr(oldIndex)
                newIndex = newIndex + 1
                oldIndex = oldIndex + 1
            End If
        Loop
        If pArrayType = BA_MULTIDIMENSION Then
            Slice = JaggedToMulti(result)
        Else
            Slice = result
        End If
    End If
    
End Function

' TODO: Implement, Document & Test FromExcelRange method
'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Stores the values contained in an Excel range in the internal array")
Public Function FromExcelRange( _
        ByRef fromRange As Object, _
        Optional ByVal columnNumber As Long = MISSING_LONG, _
        Optional ByVal rowNumber As Long = MISSING_LONG _
    ) As BetterArray
Attribute FromExcelRange.VB_Description = "Stores the values contained in an Excel range in the internal array"
    Dim localColumnNumber As Long
    Dim localRowNumber As Long
    
    If TypeName(fromRange) = "Range" Then
        If fromRange.Columns.count = 1 Then
            localColumnNumber = 1
        Else
            localColumnNumber = columnNumber
        End If
        
        If fromRange.Rows.count = 1 Then
            localRowNumber = 1
        Else
            localRowNumber = rowNumber
        End If
        
        Me.Items = fromRange.value
    
        If localColumnNumber <> MISSING_LONG Then
            If localRowNumber <> MISSING_LONG Then
                Me.Items = Me.ExtractSegment(localRowNumber, localColumnNumber)
            Else
                Me.Items = Me.ExtractSegment(, localColumnNumber)
            End If
        ElseIf localRowNumber <> MISSING_LONG Then
            Me.Items = Me.ExtractSegment(localRowNumber)
        End If
    Else
        RaiseError ErrorCodes.EC_EXPECTED_RANGE_OBJECT, "FromExcelRange()", "FromRange"
    End If
    
    Set FromExcelRange = Me
End Function

'@Ignore ProcedureNotUsed
'@Description("Extracts the segment of the array located at the specified indices.")
Public Function ExtractSegment( _
    Optional ByVal rowIndex As Long = MISSING_LONG, _
    Optional ByVal columnIndex As Long = MISSING_LONG _
    ) As Variant()
Attribute ExtractSegment.VB_Description = "Extracts the segment of the array located at the specified indices."
    Dim i As Long
    Dim localRowIndex As Long
    Dim localColumnIndex As Long
    Dim nestedBounds() As Variant
    Dim sourceArray() As Variant
    Dim result() As Variant
    
    sourceArray = GetInternalItems()

    If rowIndex = MISSING_LONG Then
        If columnIndex = MISSING_LONG Then
            ' if no coords provided return full array
            ' calls to Me.Items instead of SourceArray to covert md from jagged
            result = Me.Items
        Else
            ' Row number missing, column number present
            Select Case pArrayType
            Case BA_ONEDIMENSION
                If columnIndex >= Me.LowerBound And columnIndex <= Me.UpperBound Then
                    localColumnIndex = columnIndex
                Else
                    localColumnIndex = Me.LowerBound
                End If
                result = Array(sourceArray(localColumnIndex))
            Case BA_JAGGED, BA_MULTIDIMENSION
                ' return full column
                nestedBounds = GetMaxBoundsAtDimension(sourceArray, 2)
                If columnIndex >= nestedBounds(0) And columnIndex <= nestedBounds(1) Then
                    localColumnIndex = columnIndex
                Else
                    localColumnIndex = Me.LowerBound
                End If
                ReDim result(Me.LowerBound To Me.UpperBound)
                For i = Me.LowerBound To Me.UpperBound
                    result(i) = sourceArray(i)(localColumnIndex)
                Next
            Case BA_UNALLOCATED
                result = sourceArray
            Case Else
            End Select
        End If
    Else
        ' Row Number not missing
        If rowIndex >= Me.LowerBound And rowIndex <= Me.UpperBound Then
            localRowIndex = rowIndex
        Else
            localRowIndex = Me.LowerBound
        End If
        If columnIndex = MISSING_LONG Then
            ' Row number present, column number missing
            Select Case pArrayType
            Case BA_ONEDIMENSION
                result = Array(sourceArray(localRowIndex))
            Case BA_JAGGED, BA_MULTIDIMENSION
                result = sourceArray(localRowIndex)
            Case BA_UNALLOCATED
                result = sourceArray
            Case Else
            End Select
        Else
            ' Row number present, column number present
            Select Case pArrayType
            Case BA_ONEDIMENSION
                ' ignore column number arg as array is 1d
                result = Array(sourceArray(localRowIndex))
            Case BA_JAGGED, BA_MULTIDIMENSION
                ' return intersection of rowIndex and column number
                nestedBounds = GetMaxBoundsAtDimension(sourceArray, 2)
                If columnIndex >= nestedBounds(0) And columnIndex <= nestedBounds(1) Then
                    localColumnIndex = columnIndex
                Else
                    localColumnIndex = Me.LowerBound
                End If
                If IsArray(sourceArray(localRowIndex)(localColumnIndex)) Then
                    result = sourceArray(localRowIndex)(localColumnIndex)
                Else
                    result = Array(sourceArray(localRowIndex)(localColumnIndex))
                End If
            Case BA_UNALLOCATED
                result = sourceArray
            Case Else
            End Select
        End If
    End If
    
    ExtractSegment = result
End Function

' TODO: Implement, Document & Test ToExcelRange method
'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Writes the values stored in the array to an excel worksheet starting at the specified range.")
Public Function ToExcelRange( _
        ByRef destination As Object, _
        Optional ByVal transposeValues As Boolean _
    ) As Object
Attribute ToExcelRange.VB_Description = "Writes the values stored in the array to an excel worksheet starting at the specified range."
    Dim localRange As Object
    Dim localItems() As Variant
    Dim depth As Long
    Dim lengthRows As Long
    Dim lengthColumns As Long
    
    lengthRows = 1 'default
    lengthColumns = 1 'default
    
    If TypeName(destination) = "Range" Then
        localItems = GetInternalItems()
        depth = GetJaggedArrayDepth(localItems)
        If depth > 0 Then 'check array is allocated
            If depth = 1 Then
                '1d array
                localItems = EnsureScalar1DArray(localItems)
                If transposeValues Then
                    lengthRows = Me.length
                    localItems = Transpose_1DArray(localItems)
                Else
                    lengthColumns = Me.length
                End If
            ElseIf depth > 1 Then
                'jagged array - multidim arrays should also be stored as jagged internally
                Const outputDepth As Long = 2 'need 2 dimensions for tabular representation output
                localItems = JaggedToMulti(localItems, outputDepth, ensureScalar:=True)
                If transposeValues Then
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
            localRange.value = localItems
        Else
            'array must be unalloc
            'do nothing(?)
        End If
    Else
        RaiseError ErrorCodes.EC_EXPECTED_RANGE_OBJECT, "ToExcelRange()", "Destination"
    End If
    
    Set ToExcelRange = localRange
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Tests if the stored array is sorted in ascending order.")
Public Function IsSorted(Optional ByVal columnIndex As Long = MISSING_LONG) As Boolean
Attribute IsSorted.VB_Description = "Tests if the stored array is sorted in ascending order."
    Dim i As Long
    Dim localLowerBound As Long
    Dim localUpperBound As Long
    Dim localColumnIndex As Long
    Dim depth As Long
    Dim result As Boolean
    Dim localArray() As Variant
    Dim storedType As ArrayTypes
    
    result = True
    
    localArray = GetInternalItems()
    localLowerBound = LBound(localArray)
    localUpperBound = UBound(localArray)
    storedType = GetArrayType(localArray)
    If storedType <> BA_UNDEFINED And storedType <> BA_UNALLOCATED Then
        Select Case storedType
        Case BA_ONEDIMENSION
            For i = localLowerBound To localUpperBound - 1
                If localArray(i) > localArray(i + 1) Then
                    result = False
                    Exit For
                End If
            Next
        Case BA_MULTIDIMENSION
            ' should never be true as all md arrays stored as jagged internally
            If columnIndex = MISSING_LONG Then
                localColumnIndex = LBound(localArray, 2)
            Else
                localColumnIndex = CLng(columnIndex)
            End If
            For i = localLowerBound To localUpperBound - 1
                If i = localLowerBound Then
                End If
                If localArray(i, localColumnIndex) > localArray(i + 1, localColumnIndex) Then
                    result = False
                    Exit For
                End If
            Next
        Case BA_JAGGED
            depth = GetJaggedArrayDepth(localArray)
            If depth > 2 Then
                IsSorted = False
                RaiseError EC_EXCEEDS_MAX_SORT_DEPTH, "IsSorted"
            Else
                For i = localLowerBound To localUpperBound - 1
                    If i = localLowerBound Then
                        If columnIndex = MISSING_LONG Then
                            localColumnIndex = LBound(localArray(i))
                        Else
                            localColumnIndex = CLng(columnIndex)
                        End If
                    End If
                    If localArray(i)(localColumnIndex) > localArray(i + 1)(localColumnIndex) Then
                        result = False
                        Exit For
                    End If
                Next
            End If
        End Select
    End If
    IsSorted = result
End Function

'@Description("returns the first index at which a given element can be found in the array")
Public Function IndexOf( _
        ByVal searchElement As Variant, _
        Optional ByVal fromIndex As Long = MISSING_LONG _
    ) As Long
Attribute IndexOf.VB_Description = "returns the first index at which a given element can be found in the array"
    Dim localItems() As Variant
    Dim localLength As Long
    Dim relativeStart As Long
    Dim currentIndex As Long
    
    localItems = GetInternalItems()
    localLength = Me.length
    If localLength = 0 Then
        IndexOf = MISSING_LONG
        Exit Function
    End If

    If fromIndex = MISSING_LONG Then
        relativeStart = LBound(localItems)
    Else
        relativeStart = fromIndex
    End If
    
    If relativeStart >= LBound(localItems) Then
        currentIndex = relativeStart
    Else
        If relativeStart > 0 Then
            currentIndex = LBound(localItems)
        Else
            currentIndex = UBound(localItems) + relativeStart
        End If
        
        If currentIndex < LBound(localItems) Then
            currentIndex = LBound(localItems)
        End If
    End If
    
    Do While currentIndex <= UBound(localItems)
        If ElementsAreEqual(searchElement, localItems(currentIndex)) Then
            IndexOf = currentIndex
            Exit Function
        End If
        currentIndex = currentIndex + 1
    Loop
    
    IndexOf = MISSING_LONG
End Function

'@Description("Removes the item at the specified index. Returns the new length of the array")
Public Function Remove(ByVal index As Long) As Long
Attribute Remove.VB_Description = "Removes the item at the specified index. Returns the new length of the array"
    Dim relativeIndex As Long
    Dim localType As ArrayTypes
    
    relativeIndex = MISSING_LONG
    localType = pArrayType
    
    If index >= Me.LowerBound Then
        If index <= Me.UpperBound Then
            relativeIndex = index
        End If
    Else
        If index < 0 Then
            relativeIndex = Me.UpperBound + index
            If relativeIndex < Me.LowerBound Then relativeIndex = MISSING_LONG
        End If
    End If
    
    If relativeIndex <> MISSING_LONG Then
        Dim beforeSlice() As Variant
        Dim afterSlice() As Variant
        Dim beforeExists As Boolean
        Dim afterExists As Boolean
        
        If relativeIndex > Me.LowerBound Then
            beforeSlice = Me.Slice(Me.LowerBound, relativeIndex)
            beforeExists = True
        End If
        If relativeIndex < Me.UpperBound Then
            afterSlice = Me.Slice(relativeIndex + 1)
            afterExists = True
        End If
        If beforeExists Then
            If afterExists Then
                Dim result As BetterArray
                Set result = New BetterArray
                Me.Items = result.Concat(beforeSlice, afterSlice).Items
            Else
                Me.Items = beforeSlice
            End If
        Else
            If afterExists Then
                Me.Items = afterSlice
            End If
        End If
        If beforeExists Or afterExists Then
            If pArrayType = BA_JAGGED And localType = BA_MULTIDIMENSION Then
                pArrayType = localType
            End If
        End If
    End If
    
    Remove = Me.length
End Function


''''''''''''''''''''''''''''
' Chainable Public Methods '
''''''''''''''''''''''''''''
' TODO: Document
'@Ignore FunctionReturnValueNotUsed
'@Description("Parses a string representing an array and stores the outcome.")
Public Function ParseFromString( _
        ByVal sourceString As String, _
        Optional ByVal valueSeparator As String = ",", _
        Optional ByVal arrayOpenDelimiter As String, _
        Optional ByVal arrayClosingDelimiter As String _
    ) As BetterArray
Attribute ParseFromString.VB_Description = "Parses a string representing an array and stores the outcome."
    Dim opener As String
    Dim closer As String
    Dim ArraysAreDelimited As Boolean
    Dim result() As Variant
    
    If Len(sourceString) > 0 Then
        If arrayOpenDelimiter = vbNullString And arrayClosingDelimiter = vbNullString Then
            Dim firstChar As String
            Dim lastChar As String
            firstChar = Left$(sourceString, 1)
            lastChar = Right$(sourceString, 1)
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
            opener = arrayOpenDelimiter
            closer = arrayClosingDelimiter
            ArraysAreDelimited = True
        End If
        
        If ArraysAreDelimited Then
            result = ParseDelimitedArrayString(sourceString, valueSeparator, opener, closer)
        Else
            result = ParseArraySegmentFromString(sourceString, 1, 0, valueSeparator)
        End If
    Else
        RaiseError EC_NULL_STRING, "ParseFromString", "SourceString"
    End If
    
    Me.Items = result
    Set ParseFromString = Me
End Function

' TODO: Document & Test
'@Ignore ProcedureNotUsed
Public Function Transpose() As BetterArray
    Dim result() As Variant
    Dim localItems() As Variant
    Dim localType As ArrayTypes
    Dim storedType As ArrayTypes
    storedType = pArrayType
    localItems = GetInternalItems()
    localType = GetArrayType(localItems)
    Select Case localType
    Case ArrayTypes.BA_ONEDIMENSION
        result = Transpose_1DArray(localItems)
    Case ArrayTypes.BA_MULTIDIMENSION
        ' Sould Never match as MD arrs stored as jagged internally
        result = Transpose_2DArray(localItems)
    Case ArrayTypes.BA_JAGGED
        result = Transpose_ArrayOfArrays(localItems)
    Case Else
        result = localItems
    End Select
    Me.Items = result
    If storedType = BA_MULTIDIMENSION Then
        localType = GetArrayType(result)
        If localType = BA_JAGGED Then
            pArrayType = BA_MULTIDIMENSION
        End If
    End If
    Set Transpose = Me
End Function


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
    pItems = GetEmptyArray()
    pLength = 0
    Me.Capacity = DEFAULT_CAPACITY
    pArrayType = BA_UNALLOCATED
    Set ResetToDefault = Me
End Function

' TODO: Document & Test - Clear
'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Clears all entries in the current array but retains existing capacity.")
Public Function Clear() As BetterArray
Attribute Clear.VB_Description = "Clears all entries in the current array but retains existing capacity."
    Dim oldCapacity As Long
    oldCapacity = pCapacity
    pItems = GetEmptyArray()
    pLength = 0
    Me.Capacity = oldCapacity
    Set Clear = Me
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Joins one or more arrays onto the end of the current array.")
Public Function Concat(ParamArray args() As Variant) As BetterArray
Attribute Concat.VB_Description = "Joins one or more arrays onto the end of the current array."
    Dim arg As Variant
    Dim localArray() As Variant
    Dim localType As ArrayTypes
    Dim i As Long
    
    For Each arg In args
        If IsArray(arg) Then
            localArray = arg
            ' should never return BA_UNDEFINED as we've already checked it's an array
            localType = GetArrayType(localArray)
            ' If pushing to unallocated array short circuit individual pushes and just assign with me.items
            If pArrayType = BA_UNALLOCATED Then
                If localType = BA_UNALLOCATED Then
                    Me.Items = GetEmptyArray()
                Else
                    Me.Items = localArray
                End If
            Else
                If localType = BA_MULTIDIMENSION Then
                    localArray = MultiToJagged(localArray)
                End If
                
                If pArrayType = BA_ONEDIMENSION Then
                    ' if the stored array is not already a jagged array and
                    ' the passed array is a MD array then set our array type to MD
                    If localType = BA_MULTIDIMENSION And pArrayType <> BA_JAGGED Then
                        pArrayType = BA_MULTIDIMENSION
                    End If
                End If
                
                For i = LBound(localArray) To UBound(localArray)
                    Me.Push localArray(i)
                Next
            End If
        Else
            Me.Push arg
        End If
    Next
    Set Concat = Me
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Writes all of the elements in a Collection objet to the internal array.")
Public Function CopyFromCollection(ByVal sourceCollection As Collection) As BetterArray
Attribute CopyFromCollection.VB_Description = "Writes all of the elements in a Collection objet to the internal array."
    If sourceCollection Is Nothing Then
        RaiseError EC_EXPECTED_COLLECTION_OBJECT, "CopyFromCollection", "SourceCollection"
    End If
    Dim i As Long
    Dim newItems() As Variant
    pLength = sourceCollection.count
    If pLength = 0 Then
        newItems = GetEmptyArray
    Else
        ReDim newItems(pLowerBound To (pLength - pLowerBound - 1))
        For i = 1 To pLength
            newItems(i + pLowerBound - 1) = sourceCollection.Item(i)
        Next
    End If
    Me.Items = newItems
    Set CopyFromCollection = Me
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Sorts the stored array in ascending order.")
Public Function Sort(Optional ByVal sortColumn As Long = MISSING_LONG) As BetterArray
Attribute Sort.VB_Description = "Sorts the stored array in ascending order."
    Dim localItems() As Variant
    Dim sortedItems() As Variant
    Dim arrayType As ArrayTypes
    
    localItems = GetInternalItems()
    arrayType = GetArrayType(localItems)
    If arrayType = ArrayTypes.BA_UNALLOCATED Or _
       arrayType = ArrayTypes.BA_UNDEFINED Then
        sortedItems = GetEmptyArray
    Else
        If Me.length > 0 Then
            If arrayType <> BA_ONEDIMENSION Then
                If arrayType = ArrayTypes.BA_MULTIDIMENSION Then
                    localItems = MultiToJagged(localItems)
                End If
                Dim depth As Long
                depth = GetJaggedArrayDepth(localItems)
                If depth > 2 Then
                    RaiseError EC_EXCEEDS_MAX_SORT_DEPTH, "Sort"
                End If
            End If
            QuickSort _
                localItems, _
                LBound(localItems), _
                UBound(localItems), _
                arrayType, _
                IIf(sortColumn = MISSING_LONG, pLowerBound, sortColumn)
        End If
        sortedItems = localItems
    End If
    Me.Items = sortedItems
    Set Sort = Me
End Function

'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Copies a section of the stored array to another location in the same array.")
Public Function CopyWithin( _
        ByVal target As Long, _
        Optional ByVal startIndex As Long = MISSING_LONG, _
        Optional ByVal endIndex As Long = MISSING_LONG _
    ) As BetterArray
Attribute CopyWithin.VB_Description = "Copies a section of the stored array to another location in the same array."
    Dim localLength As Long
    Dim relativeTarget As Long
    Dim relativeStart As Long
    Dim relativeEnd As Long
    Dim toIndex As Long
    Dim fromIndex As Long
    Dim final As Long
    Dim count As Long
    Dim direction As Long
    Dim localArray() As Variant
    
    Select Case pArrayType
    Case ArrayTypes.BA_UNDEFINED
        RaiseError EC_EXPECTED_ARRAY, "CopyWithin"
    Case ArrayTypes.BA_UNALLOCATED
        RaiseError EC_UNALLOCATED_ARRAY, "CopyWithin"
    Case Else
        localArray = GetInternalItems()
        localLength = Me.length
        relativeTarget = target
        
        If relativeTarget < 0 Then
            toIndex = Max((localLength + relativeTarget), 0)
        Else
            toIndex = Min(relativeTarget, localLength)
        End If
        
        relativeStart = IIf(startIndex = MISSING_LONG, pLowerBound, startIndex)
        
        If relativeStart < 0 Then
            fromIndex = Max((localLength + relativeStart), 0)
        Else
            fromIndex = Min(relativeStart, localLength)
        End If
        
        If endIndex = MISSING_LONG Then
            relativeEnd = localLength
        Else
            relativeEnd = endIndex
        End If
        
        If relativeEnd < 0 Then
            final = Max((localLength + relativeEnd), 0)
        Else
            final = Min(relativeEnd, localLength)
        End If
        
        count = Min(final - fromIndex, localLength - toIndex)
        
        If fromIndex < toIndex And toIndex < fromIndex + count Then
            direction = -1
            fromIndex = fromIndex + count - 1
            toIndex = toIndex + count - 1
        Else
            direction = 1
        End If
        
        Do While count > 0
            If fromIndex >= LBound(localArray) And fromIndex <= UBound(localArray) Then
                localArray(toIndex) = localArray(fromIndex)
            End If
            fromIndex = fromIndex + direction
            toIndex = toIndex + direction
            count = count - 1
        Loop
        
        Me.Items = localArray
    End Select
    
    Set CopyWithin = Me

End Function

' TODO: add support for 2d arrays. Consider rename?
'@Ignore ShadowedDeclaration
'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Filters the stored array to include only values that meet the specified criteria.")
Public Function Filter( _
        ByVal match As Variant, _
        Optional ByVal include As Boolean, _
        Optional ByVal recurse As Boolean) As BetterArray
Attribute Filter.VB_Description = "Filters the stored array to include only values that meet the specified criteria."
    Dim result() As Variant
    Dim localArray() As Variant
    Dim localType As ArrayTypes
    localType = pArrayType
    If localType = BA_UNDEFINED Then
        RaiseError EC_EXPECTED_ARRAY, "Filter"
    ElseIf localType = BA_UNALLOCATED Then
        result = GetEmptyArray()
    Else
        localArray = GetInternalItems()
        result = RecursiveFilter( _
            localArray, _
            match, _
            include, _
            recurse _
        )
    End If
    Me.Items = result
    If localType = BA_MULTIDIMENSION Then pArrayType = localType
    Set Filter = Me
End Function

' TODO: Check implementation against spec
'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Reverses the order of elements in the array. The first array element becomes the last, and the last array element becomes the first.")
Public Function Reverse(Optional ByVal recurse As Boolean) As BetterArray
Attribute Reverse.VB_Description = "Reverses the order of elements in the array. The first array element becomes the last, and the last array element becomes the first."
    Dim localItems() As Variant
    Dim result() As Variant
    Dim currentType As ArrayTypes
    currentType = pArrayType
    Select Case currentType
    Case BA_UNDEFINED
        RaiseError EC_UNDEFINED_ARRAY, "Reverse"
    Case BA_ONEDIMENSION, BA_MULTIDIMENSION, BA_JAGGED
        localItems = GetInternalItems()
        result = RecursiveReverse(localItems, recurse)
        Me.Items = result
        pArrayType = currentType
    End Select
    
    Set Reverse = Me
End Function

' TODO: Implement, Document & Test Flatten method
'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("Flattens a Multi-Dimensional or Jagged array into One Dimension arrays.")
Public Function Flatten() As BetterArray
Attribute Flatten.VB_Description = "Flattens a Multi-Dimensional or Jagged array into One Dimension arrays."
    Dim localItems() As Variant
    Dim result As BetterArray
        
    ' noop on 1d or unalloc
    ' md & jagged both stored as jagged
    Select Case pArrayType
    Case BA_UNDEFINED
        RaiseError EC_UNDEFINED_ARRAY, "Flatten"
    Case BA_MULTIDIMENSION, BA_JAGGED
        localItems = GetInternalItems()
        Set result = New BetterArray
        result.LowerBound = pLowerBound
        RecursiveFlatten localItems, result
        Me.Items = result.Items
    End Select

    Set Flatten = Me
End Function

' TODO: Implement, Document & Test Shuffle method
'@Ignore ProcedureNotUsed, FunctionReturnValueNotUsed
'@Description("shuffles the order of the stored array.")
Public Function Shuffle(Optional ByVal recurse As Boolean) As BetterArray
Attribute Shuffle.VB_Description = "shuffles the order of the stored array."
    Dim localItems() As Variant
    Dim result() As Variant
    Dim currentType As ArrayTypes
    currentType = pArrayType
    Select Case currentType
    Case BA_UNDEFINED
        RaiseError EC_UNDEFINED_ARRAY, "Shuffle"
    Case BA_ONEDIMENSION, BA_MULTIDIMENSION, BA_JAGGED
        localItems = GetInternalItems()
        result = RecursiveShuffle(localItems, recurse)
        Me.Items = result
        pArrayType = currentType
    End Select
    
    Set Shuffle = Me
End Function


' TODO: Implement, Document & Test Unique
'@Description("filters the array to contain only unique elements")
Public Function Unique() As BetterArray
Attribute Unique.VB_Description = "filters the array to contain only unique elements"
    Dim localItems() As Variant
    Dim localType As ArrayTypes
    Dim result As BetterArray
    Dim i As Long
    
    localItems = GetInternalItems()
    localType = pArrayType
    Set result = New BetterArray
    For i = Me.LowerBound To Me.UpperBound
        If Me.IndexOf(localItems(i)) = i Then
            result.Push localItems(i)
        End If
    Next
    Me.Items = result.Items
    pArrayType = localType
    Set Unique = Me
End Function

' TODO: Implement, Document & Test Every method
' https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/every

' TODO: Implement, Document & Test Fill method
' https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/fill

' TODO: Implement, Document & Test Join method
' https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/join

' TODO: Implement, Document & Test lastIndexOf method
' https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/lastIndexOf

' TODO: Implement, Document & Test some method
' https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/some

' TODO: Implement, Document & Test splice method
' https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/splice

' TODO: Implement, Document & Test IncludesType method

' TODO: Implement, Document & Test FilterByType method

' TODO: Implement, Document & Test First prop ?

' TODO: Implement, Document & Test Last prop ?


'''''''''''''''''''
' Private Methods '
'''''''''''''''''''

'@Description("shuffles the order of the passed array.")
Private Function RecursiveShuffle( _
        ByRef arr() As Variant, _
        Optional ByVal recurse As Boolean _
    ) As Variant()
Attribute RecursiveShuffle.VB_Description = "shuffles the order of the passed array."
    ' https://en.wikipedia.org/wiki/Fisher%E2%80%93Yates_shuffle#The_modern_algorithm
    Dim i As Long
    Dim j As Long
    Dim lower As Long
    Dim upper As Long
    Dim nested() As Variant
    lower = LBound(arr)
    upper = UBound(arr)
    Randomize
    For i = upper To lower + 1 Step -1
        If IsArray(arr(i)) And recurse Then
            nested = arr(i)
            arr(i) = RecursiveShuffle(nested, recurse)
        End If
        j = Int(Rnd * (i + 1))
        If IsArray(arr(j)) And recurse Then
            nested = arr(j)
            arr(j) = RecursiveShuffle(nested, recurse)
        End If
        Swap arr, i, j
    Next
    RecursiveShuffle = arr
End Function

'@Description("reverses the order of the passed array.")
Private Function RecursiveReverse( _
        ByRef arr() As Variant, _
        Optional ByVal recurse As Boolean _
    ) As Variant()
Attribute RecursiveReverse.VB_Description = "reverses the order of the passed array."
    
    Dim localLength As Long
    Dim middle As Long
    Dim lower As Long
    Dim upper As Long
    Dim passThruArray() As Variant
    
    localLength = GetArrayLength(arr)
    lower = LBound(arr)
    middle = Int(localLength / 2) + lower
    
    Do While lower <> middle
        ' TODO: check this works with and lb
        upper = UBound(arr) + LBound(arr) - lower
        If IsArray(arr(lower)) And recurse Then
            passThruArray = arr(lower)
            arr(lower) = RecursiveReverse(passThruArray, recurse)
        End If
        If IsArray(arr(upper)) And recurse Then
            passThruArray = arr(upper)
            arr(upper) = RecursiveReverse(passThruArray, recurse)
        End If
        Swap arr, lower, upper
        lower = lower + 1
    Loop
    RecursiveReverse = arr
End Function

'@Description("returns the length of the passed array.")
Private Function GetArrayLength(ByRef sourceArray() As Variant) As Long
Attribute GetArrayLength.VB_Description = "returns the length of the passed array."
    GetArrayLength = UBound(sourceArray) - LBound(sourceArray) + 1
End Function

'@Description("converts all the values in a jagged array to a 1d array.")
Private Sub RecursiveFlatten( _
        ByRef arr(), _
        ByRef results As BetterArray _
    )
Attribute RecursiveFlatten.VB_Description = "converts all the values in a jagged array to a 1d array."
    Dim element As Variant
    For Each element In arr
        If IsArray(element) Then
            Dim arrayElement() As Variant
            arrayElement = element
            RecursiveFlatten arrayElement, results
        Else
            results.Push element
        End If
    Next
End Sub

'@Description("Recursively finds the largest value in an array.")
Private Function RecursiveMax(ByRef arr() As Variant) As Variant
Attribute RecursiveMax.VB_Description = "Recursively finds the largest value in an array."
    Dim i As Long
    Dim result As Variant
    
    If IsMultidimensionalArray(arr) Then
        RaiseError EC_INVALID_MULTIDIMENSIONAL_ARRAY_OPERATION, "Max", "Args"
    Else
        For i = LBound(arr) To UBound(arr)
            If IsArray(arr(i)) Then
                Dim nestedArray() As Variant
                Dim nestedResult As Variant
                nestedArray = arr(i)
                nestedResult = RecursiveMax(nestedArray)
                If IsEmpty(result) Then
                    result = nestedResult
                Else
                    If nestedResult > result Then result = nestedResult
                End If
            Else
                Dim currentValue As Variant
                currentValue = GetScalarRepresentation(arr(i))
                If currentValue <> OBJECT_REPR Then
                    If IsEmpty(result) Then
                        result = currentValue
                    Else
                        If currentValue > result Then result = currentValue
                    End If
                End If
            End If
        Next
    End If
    RecursiveMax = result
End Function

'@Description("Recursively finds the smallest value in an array.")
Private Function RecursiveMin(ByRef arr() As Variant) As Variant
Attribute RecursiveMin.VB_Description = "Recursively finds the smallest value in an array."
    Dim i As Long
    Dim result As Variant
    
    If IsMultidimensionalArray(arr) Then
        RaiseError EC_INVALID_MULTIDIMENSIONAL_ARRAY_OPERATION, "Max", "Args"
    Else
        For i = LBound(arr) To UBound(arr)
            If IsArray(arr(i)) Then
                Dim nestedArray() As Variant
                Dim nestedResult As Variant
                nestedArray = arr(i)
                nestedResult = RecursiveMin(nestedArray)
                If IsEmpty(result) Then
                    result = nestedResult
                Else
                    If nestedResult < result Then result = nestedResult
                End If
            Else
                Dim currentValue As Variant
                currentValue = GetScalarRepresentation(arr(i))
                If currentValue <> OBJECT_REPR Then
                    If IsEmpty(result) Then
                        result = currentValue
                    Else
                        If currentValue < result Then result = currentValue
                    End If
                End If
            End If
        Next
    End If
    RecursiveMin = result
End Function

'@Description("Recursively checks if the passed array includes the search element.")
Private Function RecursiveIncludes( _
        ByVal searchElement As Variant, _
        ByRef searchArray() As Variant, _
        Optional ByVal fromIndex As Long _
    ) As Boolean
Attribute RecursiveIncludes.VB_Description = "Recursively checks if the passed array includes the search element."
    Dim localLowerBound As Long
    Dim localUpperBound As Long
    Dim i As Long
    localLowerBound = LBound(searchArray)
    If fromIndex > localLowerBound Then localLowerBound = fromIndex
    localUpperBound = UBound(searchArray)
    For i = localLowerBound To localUpperBound
        If IsArray(searchArray(i)) Then
            Dim passThru() As Variant
            passThru = searchArray(i)
            If RecursiveIncludes(searchElement, passThru) Then
                RecursiveIncludes = True
                Exit Function
            End If
        Else
            If ElementsAreEqual(searchElement, searchArray(i)) Then
                RecursiveIncludes = True
                Exit Function
            End If
        End If
    Next
End Function

'@Description("Recursively filters a passed array.")
Private Function RecursiveFilter( _
        ByRef sourceArray() As Variant, _
        ByVal match As Variant, _
        ByVal include As Boolean, _
        ByVal recurse As Boolean _
    ) As Variant()
Attribute RecursiveFilter.VB_Description = "Recursively filters a passed array."
    Dim result As BetterArray
    Dim localLowerBound As Long
    Dim localUpperBound As Long
    Dim i As Long
    Dim isMatch As Boolean
    
    localLowerBound = LBound(sourceArray)
    localUpperBound = UBound(sourceArray)
    Set result = New BetterArray
    result.LowerBound = Me.LowerBound
    For i = localLowerBound To localUpperBound
        If IsArray(sourceArray(i)) Then
            If IsArrayAllocated(sourceArray(i)) Then
                If recurse Then
                    Dim localArray() As Variant
                    localArray = sourceArray(i)
                    localArray = RecursiveFilter( _
                        localArray, _
                        match, _
                        include, _
                        recurse _
                    )
                    If Not Only(localArray, Empty) Then
                        result.Push localArray
                    End If
                Else
                    If Not include Then result.Push sourceArray(i)
                End If
            Else
                result.Push GetEmptyArray()
            End If
        ElseIf IsObject(sourceArray(i)) Then
            If Not include Then result.Push sourceArray(i)
        Else
            isMatch = ElementsAreEqual(match, sourceArray(i))
            If (include And isMatch) Or (Not include And Not isMatch) Then
                result.Push sourceArray(i)
            End If
        End If
    Next
    RecursiveFilter = result.Items
End Function

'@Description("Checks if an array only contains a specific value.")
Private Function Only(ByRef arr() As Variant, ByVal match As Variant) As Boolean
Attribute Only.VB_Description = "Checks if an array only contains a specific value."
    Dim result As Boolean
    Dim element As Variant
    result = True
    For Each element In arr
        If Not ElementsAreEqual(match, element) Then
            result = False
            Exit For
        End If
    Next
    Only = result
End Function

'@Description("Compares two values for equality. Doesn't support multidimensional arrays.")
Private Function ElementsAreEqual( _
        ByVal expected As Variant, _
        ByVal actual As Variant _
    ) As Boolean
Attribute ElementsAreEqual.VB_Description = "Compares two values for equality. Doesn't support multidimensional arrays."
    ' Using 13dp of precision for EPSILON rather than IEEE 754 standard of 2^-52
    ' some roundings in type conversions cause greater diffs than machine epsilon
    Const Epsilon As Double = 0.0000000000001
    Dim result As Boolean
    Dim i As Long
    
    On Error GoTo ErrHandler
    If IsArray(expected) Or IsArray(actual) Then
        If IsArray(expected) And IsArray(actual) Then
            If LBound(expected) = LBound(actual) And _
                    UBound(expected) = UBound(actual) Then
                Dim currentlyEqual As Boolean
                currentlyEqual = True
                For i = LBound(expected) To UBound(actual)
                    If Not ElementsAreEqual(expected(i), actual(i)) Then
                        currentlyEqual = False
                        Exit For
                    End If
                Next
                result = currentlyEqual
            End If
        End If
    ElseIf IsEmpty(expected) Or IsEmpty(actual) Then
        If IsEmpty(expected) And IsEmpty(actual) Then result = True
    ElseIf IsObject(expected) Or IsObject(actual) Then
        If IsObject(expected) And IsObject(actual) Then
            If expected Is actual Then result = True
        End If
    ElseIf IsNumeric(expected) Or IsNumeric(actual) Then
        If IsNumeric(expected) And IsNumeric(actual) Then
            Dim diff As Double
            diff = Abs(expected - actual)
            If diff <= (IIf( _
                    Abs(expected) < Abs(actual), _
                    Abs(actual), _
                    Abs(expected) _
                ) * Epsilon) Then
                result = True
            End If
        End If
    ElseIf expected = actual Then
        result = True
    End If
    ElementsAreEqual = result
    Exit Function
ErrHandler:
    ElementsAreEqual = False
End Function

' TODO: Refactor to remove recursion
'@Description("Recursively parses nested arrays from a string.")
Private Function ParseDelimitedArrayString( _
        ByVal sourceString As String, _
        ByVal valueSeparator As String, _
        ByVal opener As String, _
        ByVal closer As String, _
        Optional ByRef cursor As Long = 2 _
    ) As Variant()
Attribute ParseDelimitedArrayString.VB_Description = "Recursively parses nested arrays from a string."
    Dim currentChar As String
    Dim localResult() As Variant
    Dim i As Long
    Dim breakLoop As Boolean
    Dim nextOpener As Long
    Dim nextCloser As Long
    
    currentChar = Mid$(sourceString, cursor, 1)
    nextOpener = InStr(cursor, sourceString, opener)
    nextCloser = InStr(cursor, sourceString, closer)
    
    If currentChar <> opener And (nextCloser < nextOpener Or nextOpener = 0) Then
        ' we're inside an array to be parsed
        ParseDelimitedArrayString = ParseArraySegmentFromString(sourceString, cursor, nextCloser, valueSeparator)
        Exit Function
    End If
        
    ' traverse nested arrays
    i = pLowerBound
    Do
        If currentChar = opener Then
            If i > pLowerBound Then
                ReDim Preserve localResult(i)
            Else
                localResult = GetEmptyArray
            End If
            cursor = cursor + 1
            localResult(i) = ParseDelimitedArrayString( _
                sourceString, _
                valueSeparator, _
                opener, _
                closer, _
                cursor _
            )
            i = i + 1
        End If
        currentChar = Mid$(sourceString, cursor, 1)
        Do Until currentChar = opener
            cursor = cursor + 1
            currentChar = Mid$(sourceString, cursor, 1)
            If currentChar = closer Or cursor >= Len(sourceString) Then
                breakLoop = True
                Exit Do
            End If
        Loop
    Loop Until breakLoop
    
    ParseDelimitedArrayString = localResult
End Function

'@Description("Returns an array from a delimited string")
Private Function ParseArraySegmentFromString( _
        ByVal sourceString As String, _
        ByRef cursor As Long, _
        ByVal nextCloser As Long, _
        ByVal valueSeparator As String _
    ) As Variant()
Attribute ParseArraySegmentFromString.VB_Description = "Returns an array from a delimited string"
    Dim segmentLength As Long
    Dim segment() As String
    segmentLength = IIf(nextCloser = 0, Len(sourceString), nextCloser - cursor)
    segment = Split(Mid$(sourceString, cursor, segmentLength), valueSeparator)
    cursor = nextCloser
    ParseArraySegmentFromString = DuckTypeStringArray(segment)
End Function

'@Description("Takes a string array. Returns a Variant array with the elements converted to an appropriate type")
Private Function DuckTypeStringArray(ByRef arr() As String) As Variant()
Attribute DuckTypeStringArray.VB_Description = "Takes a string array. Returns a Variant array with the elements converted to an appropriate type"
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
            Const Epsilon As Double = 2 ^ -52
            Dim diff As Double
            ' check if integer
            diff = Abs(Fix(element) - element)
            If diff > Epsilon Then
                result(i) = CDbl(element)
            Else
                result(i) = CLng(element)
            End If
        Else
            If Left$(element, 1) = Chr$(34) And Right$(element, 1) = Chr$(34) Then
                'trim string delimiters
                element = Mid$(element, 2, Len(element) - 2)
            End If
            result(i) = Trim$(element)
        End If
    Next
    DuckTypeStringArray = result
End Function

'@Description("Populates the ErrorDefinitions array.")
Private Sub PopulateErrorDefinitions()
Attribute PopulateErrorDefinitions.VB_Description = "Populates the ErrorDefinitions array."
    Dim source As String
    source = TypeName(Me)
    ErrorDefinitions(ErrorCodes.EC_EXPECTED_RANGE_OBJECT) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_EXPECTED_RANGE_OBJECT, _
        source:=source, _
        description:="Range Object Expected" _
    )
    ErrorDefinitions(ErrorCodes.EC_MAX_DIMENSIONS_LIMIT) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_MAX_DIMENSIONS_LIMIT, _
        source:=source, _
        description:="Cannot convert structure of arrays with more than 20 dimensions." _
    )
    ErrorDefinitions(ErrorCodes.EC_EXPECTED_COLLECTION_OBJECT) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_EXPECTED_COLLECTION_OBJECT, _
        source:=source, _
        description:="Valid Collection Object Expected" _
    )
    ErrorDefinitions(ErrorCodes.EC_EXCEEDS_MAX_SORT_DEPTH) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_EXCEEDS_MAX_SORT_DEPTH, _
        source:=source, _
        description:="Cannot sort on arrays with more than 2 dimensions" _
    )
    ErrorDefinitions(ErrorCodes.EC_EXPECTED_JAGGED_ARRAY) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_EXPECTED_JAGGED_ARRAY, _
        source:=source, _
        description:="Expected jagged array." _
    )
    ErrorDefinitions(ErrorCodes.EC_EXPECTED_MULTIDIMENSION_ARRAY) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_EXPECTED_MULTIDIMENSION_ARRAY, _
        source:=source, _
        description:="Expected multidimension array." _
    )
    ErrorDefinitions(ErrorCodes.EC_EXPECTED_ARRAY) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_EXPECTED_ARRAY, _
        source:=source, _
        description:="Expected array." _
    )
    ErrorDefinitions(ErrorCodes.EC_NULL_STRING) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_NULL_STRING, _
        source:=source, _
        description:="Cannot parse from a null string. Expected string with length greater than 0." _
    )
    ErrorDefinitions(ErrorCodes.EC_UNALLOCATED_ARRAY) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_UNALLOCATED_ARRAY, _
        source:=source, _
        description:="Cannot operate on unallocated array." _
    )
    ErrorDefinitions(ErrorCodes.EC_UNDEFINED_ARRAY) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_UNDEFINED_ARRAY, _
        source:=source, _
        description:="Array is undefined." _
    )
    ErrorDefinitions(ErrorCodes.EC_INVALID_MULTIDIMENSIONAL_ARRAY_OPERATION) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_INVALID_MULTIDIMENSIONAL_ARRAY_OPERATION, _
        source:=source, _
        description:="Unable to perform the requested operation on a multidimensional array." _
    )
    ErrorDefinitions(ErrorCodes.EC_EXPECTED_VARIANT_ARRAY) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_EXPECTED_VARIANT_ARRAY, _
        source:=source, _
        description:="Unable to perform the requested operation on a typed array." _
    )
End Sub

'@Description("Returns an ErrorDefinition Type populated with the provided arguments.")
Private Function ErrorDefinitionFactory( _
        ByVal number As Long, _
        ByVal source As String, _
        ByVal description As String _
    ) As ErrorDefinition
Attribute ErrorDefinitionFactory.VB_Description = "Returns an ErrorDefinition Type populated with the provided arguments."
    Dim result As ErrorDefinition
    result.number = number
    result.source = source
    result.description = description
    ErrorDefinitionFactory = result
End Function

'@Description("Raises an application error based on the passed ErrorCode and the definitions defined in ErrorDefinitions.")
Private Sub RaiseError( _
        ByVal errorCode As ErrorCodes, _
        ByVal caller As String, _
        Optional ByVal argName As String _
    )
Attribute RaiseError.VB_Description = "Raises an application error based on the passed ErrorCode and the definitions defined in ErrorDefinitions."
    Dim currentError As ErrorDefinition
    Dim localArgName As String
    If argName <> vbNullString Then localArgName = "Argument " & argName & ": "
    currentError = ErrorDefinitions(errorCode)
    Err.Raise currentError.number, _
              currentError.source & "." & caller, _
              localArgName & currentError.description
End Sub

'Used by ToExcelRange
'@Description("Ensures each element of passed array can be represented as a scalar value.")
Private Function EnsureScalar1DArray(ByRef arr() As Variant) As Variant()
Attribute EnsureScalar1DArray.VB_Description = "Ensures each element of passed array can be represented as a scalar value."
    Dim i As Long
    Dim localLowerBound As Long
    Dim localUpperBound As Long
    Dim result() As Variant
    
    localLowerBound = LBound(arr)
    localUpperBound = UBound(arr)
    ReDim result(localLowerBound To localUpperBound)
    For i = localLowerBound To localUpperBound
        result(i) = GetScalarRepresentation(arr(i))
    Next
    EnsureScalar1DArray = result
End Function

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
    ReDim result(localLowerBound To localUpperBound, _
                 localLowerBound To localLowerBound)
    j = localLowerBound
    For i = localLowerBound To localUpperBound
        result(j, localLowerBound) = arr(i)
        j = j + 1
    Next
    Transpose_1DArray = result
End Function

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
    Dim nestedBounds() As Variant
    Dim nested() As Variant
    
    nestedBounds = GetMaxBoundsAtDimension(sourceArray, 2)
    LowerBoundCol = LBound(sourceArray)
    UpperBoundCol = UBound(sourceArray)
    LowerBoundRow = nestedBounds(0)
    UpperBoundRow = nestedBounds(1)
    'check if array can fit into 1 dimension
    If LowerBoundRow = UpperBoundRow Then
        ReDim result(LowerBoundCol To UpperBoundCol)
        currentRow = LowerBoundRow
        For currentColumn = LowerBoundCol To UpperBoundCol
            result(currentColumn) = sourceArray(currentColumn)(currentRow)
        Next
    ElseIf LowerBoundCol = UpperBoundCol Then
        ReDim result(LowerBoundRow To UpperBoundRow)
        currentColumn = LowerBoundCol
        For currentRow = UpperBoundRow To UpperBoundRow
            result(currentRow) = sourceArray(currentColumn)(currentRow)
        Next
    Else
        ReDim result(LowerBoundRow To UpperBoundRow)
        For currentRow = LowerBoundRow To UpperBoundRow
            ReDim nested(LowerBoundCol To UpperBoundCol)
            For currentColumn = LowerBoundCol To UpperBoundCol
                nested(currentColumn) = sourceArray(currentColumn)(currentRow)
            Next
            result(currentRow) = nested
        Next
    End If
    Transpose_ArrayOfArrays = result
End Function

'@Description("Returns an empty 1d array with 1 slot.")
Private Function GetEmptyArray() As Variant()
Attribute GetEmptyArray.VB_Description = "Returns an empty 1d array with 1 slot."
    Dim result() As Variant
    ReDim result(pLowerBound To pLowerBound)
    GetEmptyArray = result
End Function

'@Description("Sorts in place a passed array using the QuickSort algorithm.")
Private Sub QuickSort( _
        ByRef arr() As Variant, _
        ByVal low As Long, _
        ByVal high As Long, _
        ByVal arrayType As ArrayTypes, _
        Optional ByVal col As Long _
    )
Attribute QuickSort.VB_Description = "Sorts in place a passed array using the QuickSort algorithm."
    Dim partitionIndex As Long
    If low < high Then
        partitionIndex = QsPartition(arr, low, high, arrayType, col)
        QuickSort arr, low, partitionIndex - 1, arrayType, col
        QuickSort arr, partitionIndex + 1, high, arrayType, col
    End If
End Sub

'@Description("Child function of quickSort. Returns the next partition index.")
Private Function QsPartition( _
        ByRef arr() As Variant, _
        ByVal low As Long, _
        ByVal high As Long, _
        ByVal arrayType As ArrayTypes, _
        Optional ByVal col As Long _
    ) As Long
Attribute QsPartition.VB_Description = "Child function of quickSort. Returns the next partition index."
    Dim i As Long
    Dim j As Long
    Dim pivot As Variant
    Dim current As Variant
    
    If arrayType = BA_JAGGED Then
        pivot = arr(high)(col)
    Else
        pivot = arr(high)
    End If
    
    i = low - 1
    For j = low To high - 1
        If arrayType = BA_JAGGED Then
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
    QsPartition = i + 1
End Function

'@Description("Swaps the location of two elements in an array.")
Private Sub Swap(ByRef arr() As Variant, ByVal i As Long, ByVal j As Long)
Attribute Swap.VB_Description = "Swaps the location of two elements in an array."
    Dim element As Variant
    LetOrSetElement element, arr(i)
    LetOrSetElement arr(i), arr(j)
    '@Ignore UnassignedVariableUsage
    LetOrSetElement arr(j), element
End Sub

'@Description("Sets or lets the value of a variable by reference depending on the type of the element to be assigned")
Private Sub LetOrSetElement(ByRef destination As Variant, ByRef source As Variant)
Attribute LetOrSetElement.VB_Description = "Sets or lets the value of a variable by reference depending on the type of the element to be assigned"
    If IsObject(source) Then
        Set destination = source
    Else
        destination = source
    End If
End Sub

'@Description("Retrieves the stored items trimmed to the correct length. Does not change structure (i.e jagged to multi).")
Private Function GetInternalItems() As Variant()
Attribute GetInternalItems.VB_Description = "Retrieves the stored items trimmed to the correct length. Does not change structure (i.e jagged to multi)."
    Dim result() As Variant
    If pArrayType <> BA_UNALLOCATED And pArrayType <> BA_UNDEFINED Then
        result = pItems
        If Me.Capacity > pLength Then
            If pLength > 0 Then ReDim Preserve result(pLowerBound To Me.UpperBound)
        End If
    Else
        result = GetEmptyArray()
    End If
    GetInternalItems = result
End Function

'@Description("Recursively parses an array to a string representation.")
Private Function RecursiveToString( _
        ByRef arr() As Variant, _
        ByVal prettyPrint As Boolean, _
        Optional ByVal tabs As Long = 1, _
        Optional ByVal delimitStrings As Boolean, _
        Optional ByVal openingDelimiter As String, _
        Optional ByVal closingDelimiter As String _
    ) As String
Attribute RecursiveToString.VB_Description = "Recursively parses an array to a string representation."
    Const tabWidth As Long = 2
    Dim result As String
    Dim sep As String
    Dim i As Long
    Dim opener As String
    Dim closer As String
    
    opener = IIf(openingDelimiter = vbNullString, "{", openingDelimiter)
    closer = IIf(closingDelimiter = vbNullString, "}", closingDelimiter)
    
    result = opener
    For i = LBound(arr) To UBound(arr)
        sep = IIf(i = UBound(arr), closer, IIf(prettyPrint, ", ", ","))
        If IsArray(arr(i)) Then
            Dim nested() As Variant
            nested = arr(i)
            If prettyPrint Then result = result & vbCrLf & Space(tabWidth * tabs)
            result = result & RecursiveToString(nested, prettyPrint, tabs + 1, delimitStrings, opener, closer)
            result = result & IIf(i = UBound(arr) And prettyPrint, vbCrLf & Space(tabWidth * (tabs - 1)) & sep, sep)
        Else
            result = result & Replace( _
                    CStr(GetScalarRepresentation(arr(i), delimitStrings)), _
                    ",", _
                    vbNullString _
                ) & sep
        End If
    Next
    RecursiveToString = result
End Function

'@Description("Returns scalar values as is or appropriate representation if not scalar.")
Private Function GetScalarRepresentation( _
        ByRef element As Variant, _
        Optional ByVal delimitStrings As Boolean _
    ) As Variant
Attribute GetScalarRepresentation.VB_Description = "Returns scalar values as is or appropriate representation if not scalar."
    Dim result As Variant
    If IsObject(element) Then
        On Error Resume Next
        'Try to retrieve default member
        result = element
        On Error GoTo 0
        ' if there is no default member or it returns an object result = Empty
        If IsEmpty(result) Then
            result = OBJECT_REPR
        End If
    Else
        result = element
    End If
    ' if the element is array or default member returns an array then convert tostring
    If IsArray(result) And Not IsEmpty(result) Then
        Dim passThruArray() As Variant
        passThruArray = result
        result = RecursiveToString(passThruArray, False, delimitStrings:=delimitStrings)
    End If
    If Not IsNumeric(result) And delimitStrings Then
        GetScalarRepresentation = Chr$(34) & result & Chr$(34)
    Else
        GetScalarRepresentation = result
    End If
    
End Function

'@Description("Returns the type of a passed array.")
Private Function GetArrayType(ByVal arr As Variant) As ArrayTypes
Attribute GetArrayType.VB_Description = "Returns the type of a passed array."
    Dim result As ArrayTypes
    If IsArray(arr) Then
        If Not IsArrayAllocated(arr) Then
            result = BA_UNALLOCATED
        Else
            If IsMultidimensionalArray(arr) Then
                result = BA_MULTIDIMENSION
            ElseIf IsJaggedArray(arr) Then
                result = BA_JAGGED
            Else
                result = BA_ONEDIMENSION
            End If
        End If
    Else
        If IsEmpty(arr) Then
            result = BA_UNDEFINED
        Else
            RaiseError EC_EXPECTED_ARRAY, "getArrayType", "Arr"
        End If
    End If
    GetArrayType = result
End Function

'@Description("Tests if an array has been allocated.")
Private Function IsArrayAllocated(ByVal arr As Variant) As Boolean
Attribute IsArrayAllocated.VB_Description = "Tests if an array has been allocated."
    On Error Resume Next
    IsArrayAllocated = ( _
        IsArray(arr) And _
        Not IsError(LBound(arr, 1)) And _
        LBound(arr, 1) <= UBound(arr, 1) _
    )
    On Error GoTo 0
End Function

'@Description("Tests if an array is jagged.")
Private Function IsJaggedArray(ByVal arr As Variant) As Boolean
Attribute IsJaggedArray.VB_Description = "Tests if an array is jagged."
    If IsArray(arr) Then
        On Error GoTo ErrHandler
        Dim element As Variant
        For Each element In arr
            If IsArray(element) Then
                IsJaggedArray = True
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
Private Function IsMultidimensionalArray(ByVal arr As Variant) As Boolean
Attribute IsMultidimensionalArray.VB_Description = "Tests if an array is multidimensioned."
    If IsArray(arr) Then
        On Error GoTo ErrHandler
        '@Ignore VariableNotUsed
        Dim localUpperBound As Long
        '@Ignore AssignmentNotUsed
        localUpperBound = UBound(arr, 2)
        IsMultidimensionalArray = True
        On Error GoTo 0
    End If
    Exit Function
ErrHandler:
    Err.Clear
End Function

'@Ignore FunctionReturnValueNotUsed
'@Description("re-indexes array to conform to new bounds. Does not change length.")
Private Function Rebase(Optional ByRef arr As Variant, Optional ByVal arrayType As ArrayTypes) As Variant()
Attribute Rebase.VB_Description = "re-indexes array to conform to new bounds. Does not change length."
    Dim localArray() As Variant
    Dim localType As ArrayTypes

    If IsMissing(arr) Or Not IsArray(arr) Then
        localArray = GetInternalItems
        localType = pArrayType
    Else
        localArray = arr
        If arrayType = 0 Then
            localType = GetArrayType(localArray)
        Else
            localType = arrayType
        End If
    End If
    Select Case localType
    Case ArrayTypes.BA_UNALLOCATED
        Rebase = localArray
    Case ArrayTypes.BA_MULTIDIMENSION
        ' ensure all nested arrays have the correct base
        ' if they don't and additional arrays are added via concat the resulting
        ' md array from jaggedToMulti will have unexpected bounds
        ' NOTE: This may need to be applied to jagged too for consistency
        Rebase = RecursiveRebase(localArray, True)
    Case Else
        Rebase = RecursiveRebase(localArray, False)
    End Select
End Function

'@Description("re-indexes nested arrays to conform to new bounds. Child function of Rebase)
Private Function RecursiveRebase( _
        ByRef arr() As Variant, _
        ByVal recurse As Boolean _
    ) As Variant()
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
        If IsArray(arr(i)) And recurse And offset <> 0 Then
            Dim nested() As Variant
            nested = arr(i)
            newItems(i + offset) = RecursiveRebase(nested, recurse)
        Else
            newItems(i + offset) = arr(i)
        End If
    Next
    
    RecursiveRebase = newItems
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

'@Description("Converts jagged arrays to multidimension arrays")
Private Function JaggedToMulti( _
        ByRef sourceArray() As Variant, _
        Optional ByVal depth As Long, _
        Optional ByVal ensureScalar As Boolean _
    ) As Variant()
Attribute JaggedToMulti.VB_Description = "Converts jagged arrays to multidimension arrays"
    Dim localDepth As Long
    Dim result() As Variant
    localDepth = depth
    If localDepth = 0 Then
        localDepth = GetJaggedArrayDepth(sourceArray)
'        If Not isJaggedArray(arr) Then
'            RaiseError EC_EXPECTED_JAGGED_ARRAY, "jaggedToMulti", "Arr()"
'            Exit Function
'        End If
'        localDepth = getJaggedArrayDepth(arr)
    End If
    If localDepth <= 1 Then
        result = sourceArray
    ElseIf localDepth = 2 Then
        Dim firstDimBounds() As Variant
        Dim secondDimBounds() As Variant
        Dim i As Long
        Dim j As Long
        firstDimBounds = GetArrayBounds(sourceArray)
        secondDimBounds = GetMaxBoundsAtDimension(sourceArray, 2)
        ReDim result(firstDimBounds(0) To firstDimBounds(1), _
            secondDimBounds(0) To secondDimBounds(1))
        For i = firstDimBounds(0) To firstDimBounds(1)
            If IsArray(sourceArray(i)) Then
                For j = LBound(sourceArray(i)) To UBound(sourceArray(i))
                    If ensureScalar Then
                        LetOrSetElement result(i, j), _
                            GetScalarRepresentation(sourceArray(i)(j))
                    Else
                        LetOrSetElement result(i, j), _
                            sourceArray(i)(j)
                    End If
                Next
            Else
                If ensureScalar Then
                    LetOrSetElement result(i, LBound(result, 2)), _
                        GetScalarRepresentation(sourceArray(i))
                Else
                    LetOrSetElement result(i, LBound(result, 2)), _
                        sourceArray(i)
                End If
            End If
        Next
    ElseIf localDepth > 2 And localDepth <= 20 Then
        result = RecursiveJaggedToMulti(sourceArray, localDepth, ensureScalar:=localDepth)
    Else
        RaiseError EC_MAX_DIMENSIONS_LIMIT, "jaggedToMulti()", "SourceArray()"
        result = sourceArray
    End If
    JaggedToMulti = result
End Function


'@Description("Recursively Converts jagged arrays to multidimension arrays")
Private Function RecursiveJaggedToMulti( _
        ByRef arr() As Variant, _
        Optional ByVal depth As Long, _
        Optional ByVal currentDepth As Long, _
        Optional ByRef crumbs As Variant, _
        Optional ByRef result As Variant, _
        Optional ByVal ensureScalar As Boolean _
    ) As Variant()
Attribute RecursiveJaggedToMulti.VB_Description = "Recursively Converts jagged arrays to multidimension arrays"
    Dim i As Long
    Dim localLowerBound As Long
    Dim localUpperBound As Long
    Dim localDepth As Long
    Dim localCurrentDepth As Long
    Dim localCrumbs() As Variant
    Dim localResult() As Variant


    localDepth = depth
    localCurrentDepth = currentDepth

    If IsMissing(crumbs) Then
        ReDim localCrumbs(localDepth - 1)
    Else
        localCrumbs = crumbs
    End If

    If IsMissing(result) Then
        localResult = CreateMultidimensionalArray(MapJaggedArray(arr, KnownDepth:=localDepth))
    Else
        localResult = result
    End If

    localCurrentDepth = localCurrentDepth + 1
    localLowerBound = LBound(arr)
    localUpperBound = UBound(arr)
    For i = localLowerBound To localUpperBound
        localCrumbs(localCurrentDepth - 1) = i
        If localCurrentDepth = localDepth Then
            If ensureScalar Then
                localResult = LetElementByBreadcrumb(localResult, localCrumbs, GetScalarRepresentation(arr(i)))
            Else
                localResult = LetElementByBreadcrumb(localResult, localCrumbs, arr(i))
            End If
        Else
            Dim nested() As Variant
            If Not IsArray(arr(i)) Then
                nested = Array(arr(i))
            Else
                nested = arr(i)
            End If
            localResult = RecursiveJaggedToMulti( _
                arr:=nested, _
                depth:=localDepth, _
                currentDepth:=localCurrentDepth, _
                crumbs:=localCrumbs, _
                result:=localResult, _
                ensureScalar:=ensureScalar _
            )
        End If
    Next
    RecursiveJaggedToMulti = localResult
End Function

'@Description("Converts multidimension arrays to jagged arrays.")
Private Function MultiToJagged( _
        ByRef arr() As Variant, _
        Optional ByVal depth As Long, _
        Optional ByVal currentDepth As Long, _
        Optional ByRef crumbs As Variant _
    ) As Variant()
Attribute MultiToJagged.VB_Description = "Converts multidimension arrays to jagged arrays."
    Dim i As Long
    Dim localLowerBound As Long
    Dim localUpperBound As Long
    Dim localDepth As Long
    Dim localCurrentDepth As Long
    Dim localCrumbs() As Variant
    Dim result() As Variant
    Dim boundOffset As Long
    
    localDepth = depth
    localCurrentDepth = currentDepth
    
    If localDepth = 0 Then
        If Not IsMultidimensionalArray(arr) Then
            RaiseError EC_EXPECTED_MULTIDIMENSION_ARRAY, "multiToJagged", "Arr()"
            Exit Function
        End If
        localDepth = GetMultidimensionalArrayDepth(arr)
    End If
    
    If localDepth > 20 Then
        RaiseError EC_MAX_DIMENSIONS_LIMIT, "multiToJagged", "Arr()"
    ElseIf localDepth > 1 Then
        If IsMissing(crumbs) Then
            ReDim localCrumbs(localDepth - 1)
        Else
            localCrumbs = crumbs
        End If
        localCurrentDepth = localCurrentDepth + 1
        localLowerBound = LBound(arr, localCurrentDepth)
        localUpperBound = UBound(arr, localCurrentDepth)
        boundOffset = pLowerBound - localLowerBound
        ReDim result(localLowerBound + boundOffset To localUpperBound + boundOffset)
        For i = localLowerBound To localUpperBound
            localCrumbs(localCurrentDepth - 1) = i
            If localCurrentDepth = localDepth Then
                LetOrSetElement result(i + boundOffset), GetElementByBreadcrumb(arr, localCrumbs)
            Else
                result(i + boundOffset) = MultiToJagged(arr, localDepth, localCurrentDepth, localCrumbs)
            End If
        Next
    Else
        result = arr
    End If
    MultiToJagged = result
End Function

'@Description("Returns a number representing how many levels of nested arrays are contained in array.")
Private Function GetJaggedArrayDepth(ByRef arr() As Variant) As Long
Attribute GetJaggedArrayDepth.VB_Description = "Returns a number representing how many levels of nested arrays are contained in array."
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
                currentDepth = GetJaggedArrayDepth(nested)
                If currentDepth > maxDepth Then maxDepth = currentDepth
            End If
        Next
        depth = depth + maxDepth
    End If
    GetJaggedArrayDepth = depth
End Function

'@Description("Gets the maximium bound values for a specified dimension of a jagged array.")
Private Function GetMaxBoundsAtDimension( _
        ByRef sourceArray() As Variant, _
        Optional ByVal dimension As Long = 1, _
        Optional ByVal currentDimension As Long = 1 _
    ) As Variant()
Attribute GetMaxBoundsAtDimension.VB_Description = "Gets the maximium bound values for a specified dimension of a jagged array."
    Dim maxBounds() As Variant
    
    If currentDimension < dimension Then
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
                currentBounds = GetMaxBoundsAtDimension(nested, dimension, currentDimension + 1)
                If Not IsArrayAllocated(maxBounds) Then
                    maxBounds = currentBounds
                Else
                    If currentBounds(0) < maxBounds(0) Then maxBounds(0) = currentBounds(0)
                    If currentBounds(1) > maxBounds(1) Then maxBounds(1) = currentBounds(1)
                End If
            End If
        Next
    Else
        maxBounds = GetArrayBounds(sourceArray)
    End If
    
    GetMaxBoundsAtDimension = maxBounds
End Function

'@Description("Returns an array representing the structure of a jagged array")
Private Function MapJaggedArray( _
        ByRef sourceArray() As Variant, _
        Optional ByVal KnownDepth As Long _
    ) As Variant()
Attribute MapJaggedArray.VB_Description = "Returns an array representing the structure of a jagged array"
    Dim i As Long
    Dim localMap() As Variant
    Dim localDepth As Long

    If Not IsArray(sourceArray) Then Exit Function

    If KnownDepth > 0 Then
        localDepth = KnownDepth
    Else
        localDepth = GetJaggedArrayDepth(sourceArray)
    End If

    ' crumbs are base 0
    ReDim localMap(0 To localDepth - 1)

    For i = LBound(localMap) To UBound(localMap)
        localMap(i) = GetMaxBoundsAtDimension(sourceArray, i + 1)
    Next

    MapJaggedArray = localMap
End Function

'@Description("Returns a 0-based array with two slots. First slot is the lowerbound of the passed array. second slot (index 0) is the upper bound.")
Private Function GetArrayBounds(ByVal arr As Variant) As Variant()
Attribute GetArrayBounds.VB_Description = "Returns a 0-based array with two slots. First slot is the lowerbound of the passed array. second slot (index 0) is the upper bound."
    Dim result(0 To 1) As Variant
    If IsArray(arr) Then
        result(0) = LBound(arr)
        result(1) = UBound(arr)
    End If
    GetArrayBounds = result
End Function

'@Description("Returns a number representing the total number of dimensions of a multidimension array.")
Private Function GetMultidimensionalArrayDepth(ByVal arr As Variant) As Long
Attribute GetMultidimensionalArrayDepth.VB_Description = "Returns a number representing the total number of dimensions of a multidimension array."
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
    GetMultidimensionalArrayDepth = i - 1
End Function

'@Description("Returns the element stored in an array based on a crumb argument. A crumb is an array representing the location of the multidimension array slot.")
Private Function GetElementByBreadcrumb( _
        ByRef arr() As Variant, _
        ByRef crumb() As Variant _
    ) As Variant
Attribute GetElementByBreadcrumb.VB_Description = "Returns the element stored in an array based on a crumb argument. A crumb is an array representing the location of the multidimension array slot."
    Dim result As Variant
    Select Case UBound(crumb)
    Case 0
        LetOrSetElement result, arr(crumb(0))
    Case 1
        LetOrSetElement result, arr(crumb(0), crumb(1))
    Case 2
        LetOrSetElement result, arr(crumb(0), crumb(1), crumb(2))
    Case 3
        LetOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3))
    Case 4
        LetOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4))
    Case 5
        LetOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5))
    Case 6
        LetOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6))
    Case 7
        LetOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7))
    Case 8
        LetOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8))
    Case 9
        LetOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9))
    Case 10
        LetOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10))
    Case 11
        LetOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10), crumb(11))
    Case 12
        LetOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10), crumb(11), crumb(12))
    Case 13
        LetOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10), crumb(11), crumb(12), crumb(13))
    Case 14
        LetOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10), crumb(11), crumb(12), crumb(13), crumb(14))
    Case 15
        LetOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15))
    Case 16
        LetOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16))
    Case 17
        LetOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17))
    Case 18
        LetOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18))
    Case 19
        LetOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18), crumb(19))
    Case 20
        LetOrSetElement result, arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18), crumb(19), crumb(20))
    End Select
    '@Ignore UnassignedVariableUsage
    GetElementByBreadcrumb = result
End Function

'@Description("Assigns an element to a location in a multidimensonal array based on a crumb argument. A crumb is an array representing the location of the multidimension array slot.")
Private Function LetElementByBreadcrumb( _
        ByRef arr() As Variant, _
        ByRef crumb() As Variant, _
        ByVal element As Variant _
    ) As Variant
Attribute LetElementByBreadcrumb.VB_Description = "Assigns an element to a location in a multidimensonal array based on a crumb argument. A crumb is an array representing the location of the multidimension array slot."
    Select Case UBound(crumb)
    Case 0
        LetOrSetElement arr(crumb(0)), element
    Case 1
        LetOrSetElement arr(crumb(0), crumb(1)), element
    Case 2
        LetOrSetElement arr(crumb(0), crumb(1), crumb(2)), element
    Case 3
        LetOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3)), element
    Case 4
        LetOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4)), element
    Case 5
        LetOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5)), element
    Case 6
        LetOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6)), element
    Case 7
        LetOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7)), element
    Case 8
        LetOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8)), element
    Case 9
        LetOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9)), _
            element
    Case 10
        LetOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
            crumb(10)), element
    Case 11
        LetOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
            crumb(10), crumb(11)), element
    Case 12
        LetOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
            crumb(10), crumb(11), crumb(12)), element
    Case 13
        LetOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
            crumb(10), crumb(11), crumb(12), crumb(13)), element
    Case 14
        LetOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
            crumb(10), crumb(11), crumb(12), crumb(13), crumb(14)), element
    Case 15
        LetOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
            crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15)), element
    Case 16
        LetOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
        crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16)), element
    Case 17
        LetOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
            crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17)), element
    Case 18
        LetOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
            crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18)), element
    Case 19
        LetOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
        crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18), crumb(19)), element
    Case 20
        LetOrSetElement arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
        crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18), crumb(19), crumb(20)), _
        element
    End Select

    LetElementByBreadcrumb = arr
End Function

'@Description("Returns an empty multidimensional array with the dimensions specified in the crumb argument. A crumb is an array representing the structure of the multidimension array.")
Private Function CreateMultidimensionalArray(ByRef crumb() As Variant) As Variant()
Attribute CreateMultidimensionalArray.VB_Description = "Returns an empty multidimensional array with the dimensions specified in the crumb argument. A crumb is an array representing the structure of the multidimension array."
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
    CreateMultidimensionalArray = result
End Function


























