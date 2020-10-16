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
' Copyright (c) 2019 Nick Morgan. See License: https://github.com/Senipah/VBA-Better-Array/blob/master/LICENSE
'@IgnoreModule ProcedureNotUsed
'@IgnoreModule FunctionReturnValueNotUsed
'@IgnoreModule FunctionReturnValueDiscarded
'@IgnoreModule FunctionReturnValueAlwaysDiscarded
Option Explicit

'''''''''''''''''''
' Constant Fields '
'''''''''''''''''''

Private Const DEFAULT_CAPACITY As Long = 4
' NOTE: placeholder
Private Const MAX_ARRAY_LENGTH As Long = &H7FEFFFFF
Private Const OBJECT_REPR As String = "OBJECT"
Private Const MISSING_LONG As Long = -9999

Public Enum ArrayTypes
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
    EC_EXCEEDS_MAX_ARRAY_LENGTH
    EC_STRING_TYPE_EXPECTED
    EC_CANNOT_CONVERT_TO_REQUESTED_STRUCTURE
    EC_END
End Enum

Public Enum ComparisonType
    CT_EQUALITY
    CT_LIKENESS
End Enum

Private Type ErrorDefinition
    number As Long
    Source As String
    description As String
End Type


''''''''''
' Fields '
''''''''''

Private Type TFields
    Capacity As Long
    Length As Long
    LowerBound As Long
    Items() As Variant
    ArrayType As ArrayTypes
    ErrorDefinitions(EC_START To EC_END) As ErrorDefinition
    LowerBoundSet As Boolean
End Type

Private Type TString
    Text As String
    Length As Long
    ByteLength As Long
End Type

Private this As TFields

'@Description("Constructor")
Private Sub Class_Initialize()
Attribute Class_Initialize.VB_Description = "Constructor"
    Me.Capacity = DEFAULT_CAPACITY
    this.ArrayType = BA_UNALLOCATED
    PopulateErrorDefinitions
End Sub

'''''''''''''''''''''
' Public Properties '
'''''''''''''''''''''

'@Description("Returns the capacity of the internal array.")
Public Property Get Capacity() As Long
Attribute Capacity.VB_Description = "Returns the capacity of the internal array."
    Capacity = this.Capacity
End Property

'@Description("Sets the capacity of the internal array.")
'@Ignore ProcedureNotUsed
Public Property Let Capacity(ByVal value As Long)
Attribute Capacity.VB_Description = "Sets the capacity of the internal array."
    '@Ignore FunctionReturnValueDiscarded
    If value < 0 Then Err.Raise 9
    If value <> this.Capacity Then
        If this.Capacity > 0 Then
            If GetArrayLength(this.Items) <> value Then
                Dim newItems() As Variant
                newItems = this.Items
                ReDim Preserve newItems(this.LowerBound To (value + this.LowerBound - 1))
                InternalItems = newItems
            End If
        Else
            ReDim this.Items(this.LowerBound To (DEFAULT_CAPACITY + this.LowerBound - 1))
        End If
        this.Capacity = UBound(this.Items) - this.LowerBound + 1
    End If
End Property

'@Description("Returns the number of entries stored in the array.")
Public Property Get Length() As Long
Attribute Length.VB_Description = "Returns the number of entries stored in the array."
    Length = this.Length
End Property

'@Description("Returns largest used index in the array.")
Public Property Get UpperBound() As Long
Attribute UpperBound.VB_Description = "Returns largest used index in the array."
    If this.ArrayType = BA_UNALLOCATED Then
        UpperBound = -1
    Else
        UpperBound = this.Length + this.LowerBound - 1
    End If
End Property

'@Description("Returns smallest used index in the array.")
Public Property Get LowerBound() As Long
Attribute LowerBound.VB_Description = "Returns smallest used index in the array."
    LowerBound = this.LowerBound
End Property

'@Description("Sets the starting index of the array. If an array is already stored then it will be re-indexed to match the new value")
Public Property Let LowerBound(ByVal value As Long)
Attribute LowerBound.VB_Description = "Sets the starting index of the array. If an array is already stored then it will be re-indexed to match the new value"
    this.LowerBoundSet = True
    If value <> this.LowerBound Then
        this.LowerBound = value
        InternalItems = Rebase()
        this.Capacity = GetArrayLength(this.Items)
    End If
End Property

'@DefaultMember
'@Description("Returns the element in the array stored at the specified index.")
Public Property Get Item(ByVal index As Long) As Variant
Attribute Item.VB_Description = "Returns the element in the array stored at the specified index."
Attribute Item.VB_UserMemId = 0
    If index <= this.Length Then
        ' not using letOrSetElement here as it might try to access the Items letter
        If IsObject(this.Items(index)) Then
            Set Item = this.Items(index)
        Else
            Item = this.Items(index)
        End If
    Else
        '@Ignore FunctionReturnValueDiscarded
        Err.Raise 9
    End If
End Property

'@Description("Assigns the passed element to the array at the specified index.")
Public Property Let Item(ByVal index As Long, ByVal element As Variant)
Attribute Item.VB_Description = "Assigns the passed element to the array at the specified index."
    If Me.UpperBound >= index Then
        If index < this.LowerBound Then
            '@Ignore FunctionReturnValueDiscarded
            Me.Unshift element
        Else
            LetOrSetElement this.Items(index), element
        End If
    Else
        '@Ignore FunctionReturnValueDiscarded
        Me.Push element
    End If
End Property

'@Description("Returns the stored array.")
Public Property Get Items() As Variant
Attribute Items.VB_Description = "Returns the stored array."
    Dim result() As Variant
    result = InternalItems
    If this.ArrayType = ArrayTypes.BA_MULTIDIMENSION Then
        If IsJaggedArray(result) Then result = JaggedToMulti(result)
    End If
    Items = result
End Property

'@Description("Stores the passed array.")
Public Property Let Items(ByVal values As Variant)
Attribute Items.VB_Description = "Stores the passed array."
    Dim localLowerBound As Long
    Dim localValues() As Variant
    
    If TypeName(values) = TypeName(Me) Then
        localValues = values.Items
    ElseIf IsArray(values) Then
        If TypeName(values) <> "Variant()" Then
            RaiseError EC_EXPECTED_VARIANT_ARRAY, "Items", "Values"
            Exit Property
        Else
            localValues = values
        End If
    Else
        ' If Values is not an array then clear our internal array and push the Values arg
        ' NOTE: this possibly violates PoLA - assess for refactor
        '@Ignore FunctionReturnValueDiscarded
        Me.Clear.Push values
        Exit Property
    End If

    this.ArrayType = GetArrayType(localValues)
    
    If this.ArrayType = ArrayTypes.BA_UNDEFINED Then
        RaiseError EC_EXPECTED_ARRAY, "Items", "Values"
    Else
        If this.ArrayType = BA_UNALLOCATED Then
            localValues = GetEmptyArray
        End If
        
        localLowerBound = LBound(localValues)
        
        If Not this.LowerBoundSet Then
            this.LowerBound = localLowerBound
        End If
        
        If this.ArrayType = BA_MULTIDIMENSION Then
            localValues = MultiToJagged(localValues)
        End If
        
        If localLowerBound <> this.LowerBound Then
            InternalItems = Rebase(localValues, this.ArrayType)
        Else
            InternalItems = localValues
        End If
        this.Length = GetArrayLength(localValues)
        Me.Capacity = this.Length
    End If

End Property

Public Property Get ArrayType() As ArrayTypes
    ArrayType = this.ArrayType
End Property

Public Property Let ArrayType(ByVal NewType As ArrayTypes)
    Select Case NewType
    Case ArrayTypes.BA_UNDEFINED
        RaiseError EC_CANNOT_CONVERT_TO_REQUESTED_STRUCTURE, "ArrayType", "NewType"
    Case ArrayTypes.BA_UNALLOCATED
        If this.ArrayType <> BA_UNALLOCATED Then
            Me.ResetToDefault
        End If
    Case ArrayTypes.BA_ONEDIMENSION
        Select Case this.ArrayType
        Case ArrayTypes.BA_MULTIDIMENSION, ArrayTypes.BA_JAGGED
            RaiseError EC_CANNOT_CONVERT_TO_REQUESTED_STRUCTURE, "ArrayType", "NewType"
        Case Else
            this.ArrayType = NewType
        End Select
    Case Else
        this.ArrayType = NewType
    End Select
End Property


''''''''''''''''''''''
' Private Properties '
''''''''''''''''''''''

'@Description("Retrieves the stored items trimmed to the correct length. Does not change structure (i.e jagged to multi).")
Private Property Get InternalItems() As Variant()
Attribute InternalItems.VB_Description = "Retrieves the stored items trimmed to the correct length. Does not change structure (i.e jagged to multi)."
    Dim result() As Variant
    If this.ArrayType <> BA_UNALLOCATED And this.ArrayType <> BA_UNDEFINED Then
        result = this.Items
        If this.Capacity > this.Length Then
            If this.Length = 0 Then
                ReDim Preserve result(this.LowerBound To this.LowerBound)
            ElseIf this.Length > 0 Then
                ReDim Preserve result(this.LowerBound To Me.UpperBound)
            End If
        End If
    Else
        result = GetEmptyArray
    End If
    InternalItems = result
End Property

'@Description("Assigns the passed array to this.items")
Private Property Let InternalItems(ByRef value() As Variant)
Attribute InternalItems.VB_Description = "Assigns the passed array to this.items"
    this.Items = value
End Property


''''''''''''''''''
' Public Methods '
''''''''''''''''''

'@Description("Adds the provided argments to the array at the next available index.")
Public Function Push(ParamArray args() As Variant) As Variant
Attribute Push.VB_Description = "Adds the provided argments to the array at the next available index."
    Dim element As Variant

    If this.ArrayType = ArrayTypes.BA_UNALLOCATED Or _
       this.ArrayType = ArrayTypes.BA_UNDEFINED Then
        this.ArrayType = ArrayTypes.BA_ONEDIMENSION
    End If

    For Each element In args
        If this.Length = this.Capacity Then
            EnsureCapacity this.Length + 1
        End If
        If IsArray(element) Then
            Dim arrayElement() As Variant
            Dim ArrayElementType As ArrayTypes
            arrayElement = element
            If this.ArrayType = BA_ONEDIMENSION Then
                ArrayElementType = GetArrayType(arrayElement)
                If ArrayElementType = BA_MULTIDIMENSION Then
                    this.ArrayType = BA_MULTIDIMENSION
                Else
                    this.ArrayType = BA_JAGGED
                End If
            End If
            If LBound(arrayElement) <> this.LowerBound Then
                arrayElement = Rebase(arrayElement, ArrayElementType)
            End If
            LetOrSetElement this.Items(this.Length + this.LowerBound), arrayElement
        Else
            LetOrSetElement this.Items(this.Length + this.LowerBound), element
        End If
        this.Length = this.Length + 1
    Next
    Push = this.Length
End Function

'@Description("Removes the last element from the array and returns that element.")
Public Function Pop() As Variant
Attribute Pop.VB_Description = "Removes the last element from the array and returns that element."
    Dim result As Variant
    Dim newItems() As Variant
    If this.Length > 0 Then
        result = this.Items(Me.UpperBound)
        newItems = Me.Slice(this.LowerBound, Me.UpperBound)
        Me.Items = newItems
    End If
    Pop = result
End Function

'@Description("Removes the first element from the array and returns that element.")
Public Function Shift() As Variant
Attribute Shift.VB_Description = "Removes the first element from the array and returns that element."
    Dim newItems() As Variant
    Dim result As Variant
    If this.Length > 0 Then
        result = this.Items(this.LowerBound)
        newItems = Me.Slice(this.LowerBound + 1)
        Me.Items = newItems
    End If
    Shift = result
End Function

'@Description("Adds one or more elements to the beginning of an array. Returns the new length of the array.")
Public Function Unshift(ParamArray args() As Variant) As Long
Attribute Unshift.VB_Description = "Adds one or more elements to the beginning of an array. Returns the new length of the array."
    Dim newItems() As Variant
    Dim oldItems() As Variant
    Dim oldType As ArrayTypes
    
    newItems = args
    oldType = this.ArrayType
    oldItems = InternalItems
    Me.Items = newItems
    If oldType <> BA_UNALLOCATED And oldType <> BA_UNDEFINED Then
        '@Ignore FunctionReturnValueDiscarded
        Me.Concat oldItems
        this.ArrayType = oldType
    End If
    Unshift = this.Length
End Function

'@Description("Returns a string representing the array structure and its elements.")
Public Function ToString( _
        Optional ByVal prettyPrint As Boolean, _
        Optional ByVal separator As String = ",", _
        Optional ByVal openingDelimiter As String = "{", _
        Optional ByVal closingDelimiter As String = "}", _
        Optional ByVal quoteStrings As Boolean _
    ) As String
Attribute ToString.VB_Description = "Returns a string representing the array structure and its elements."
    Dim localArrayType As ArrayTypes
    Dim result As String
    Dim localItems() As Variant
    Dim sep As String
    
    sep = IIf(prettyPrint, separator & Space(1), separator)
    
    localItems = InternalItems
    localArrayType = GetArrayType(localItems)
    If Not localArrayType = ArrayTypes.BA_UNDEFINED And _
       Not localArrayType = ArrayTypes.BA_UNALLOCATED Then
        If localArrayType = ArrayTypes.BA_MULTIDIMENSION Then
            localItems = MultiToJagged(localItems)
        End If
        RecursiveToString _
            sourceArray:=localItems, _
            prettyPrint:=prettyPrint, _
            separator:=sep, _
            openingDelimiter:=openingDelimiter, _
            closingDelimiter:=closingDelimiter, _
            quoteStrings:=quoteStrings
        result = StringBuilder(Final:=True)
    End If
    ToString = result
End Function

'@Description("Determines whether the array includes a certain value among its entries.")
Public Function Includes( _
        ByVal searchElement As Variant, _
        Optional ByVal fromIndex As Long = MISSING_LONG, _
        Optional ByVal recurse As Boolean _
    ) As Boolean
Attribute Includes.VB_Description = "Determines whether the array includes a certain value among its entries."
    Dim localLength As Long
    Dim currentIndex As Long
    Dim searchArray() As Variant
    
    searchArray = InternalItems
    localLength = this.Length
    
    If localLength = 0 Then
        Includes = False
        Exit Function
    End If
        
    If fromIndex > this.LowerBound Then
        currentIndex = fromIndex
    ElseIf fromIndex = MISSING_LONG Then
        currentIndex = this.LowerBound
    Else
        currentIndex = localLength + fromIndex
    End If
    
    Includes = RecursiveIncludes(searchElement, searchArray, currentIndex, recurse:=recurse)
End Function

'@Description("Determines whether the type of any of the elements in the array matches the typename.")
Public Function IncludesType( _
        ByVal searchTypeName As String, _
        Optional ByVal fromIndex As Long = MISSING_LONG, _
        Optional ByVal recurse As Boolean _
    ) As Boolean
Attribute IncludesType.VB_Description = "Determines whether the type of any of the elements in the array matches the typename."
    Dim localLength As Long
    Dim currentIndex As Long
    Dim searchArray() As Variant
    
    searchArray = InternalItems
    localLength = this.Length
    
    If localLength = 0 Then
        IncludesType = False
        Exit Function
    End If
        
    If fromIndex > this.LowerBound Then
        currentIndex = fromIndex
    ElseIf fromIndex = MISSING_LONG Then
        currentIndex = this.LowerBound
    Else
        currentIndex = localLength + fromIndex
    End If
    
    IncludesType = RecursiveIncludes(searchTypeName, searchArray, currentIndex, True, recurse)
End Function

'@Description("Determines whether all elements in the array match a value.")
Public Function Every( _
        ByVal searchElement As Variant, _
        Optional ByVal fromIndex As Long = MISSING_LONG _
    ) As Boolean
Attribute Every.VB_Description = "Determines whether all elements in the array match a value."
    
    Dim localLength As Long
    Dim currentIndex As Long
    Dim searchArray() As Variant
    
    searchArray = InternalItems
    localLength = this.Length
    
    If localLength = 0 Then
        Every = False
        Exit Function
    End If
        
    If fromIndex > this.LowerBound Then
        currentIndex = fromIndex
    ElseIf fromIndex = MISSING_LONG Then
        currentIndex = this.LowerBound
    Else
        currentIndex = localLength + fromIndex
    End If
    
    Every = RecursiveEvery(searchElement, searchArray, currentIndex)
End Function

'@Description("Determines whether the type of all elements match the provided typename.")
Public Function EveryType( _
        ByVal searchTypeName As String, _
        Optional ByVal fromIndex As Long = MISSING_LONG _
    ) As Boolean
Attribute EveryType.VB_Description = "Determines whether the type of all elements match the provided typename."
    
    Dim localLength As Long
    Dim currentIndex As Long
    Dim searchArray() As Variant
    
    searchArray = InternalItems
    localLength = this.Length
    
    If localLength = 0 Then
        EveryType = False
        Exit Function
    End If
        
    If fromIndex > this.LowerBound Then
        currentIndex = fromIndex
    ElseIf fromIndex = MISSING_LONG Then
        currentIndex = this.LowerBound
    Else
        currentIndex = localLength + fromIndex
    End If
    
    EveryType = RecursiveEvery(searchTypeName, searchArray, currentIndex, True)
    
End Function

'@Description("Returns a new array that contains the keys for each index in the array.")
Public Function Keys() As Variant()
Attribute Keys.VB_Description = "Returns a new array that contains the keys for each index in the array."
    Dim i As Long
    Dim localLowerBound As Long
    Dim result() As Variant
    
    If this.ArrayType = BA_UNDEFINED Then
        RaiseError EC_UNDEFINED_ARRAY, "Keys"
    ElseIf this.ArrayType = BA_UNALLOCATED Then
        RaiseError EC_UNALLOCATED_ARRAY, "Keys"
    Else
        localLowerBound = this.LowerBound
        ReDim result(0 To this.Length - 1)
        For i = LBound(result) To UBound(result)
            result(i) = i + localLowerBound
        Next
    End If
    Keys = result
End Function

'@Description("Returns the largest value in a list of values. If arguments are provided it will return the largest argument. If no arguments are provided it will return the largest element in the stored array")
Public Function Max(ParamArray args() As Variant) As Variant
Attribute Max.VB_Description = "Returns the largest value in a list of values. If arguments are provided it will return the largest argument. If no arguments are provided it will return the largest element in the stored array"
    Dim localItems() As Variant
    
    If UBound(args) < LBound(args) Then          'Test if no args provided
        localItems = InternalItems
    Else
        localItems = args
    End If
    
    Max = RecursiveMax(localItems)
End Function

'@Description("Returns the smallest value in a list of values. If arguments are provided it will return the smallest argument. If no arguments are provided it will return the smallest element in the stored array")
Public Function Min(ParamArray args() As Variant) As Variant
Attribute Min.VB_Description = "Returns the smallest value in a list of values. If arguments are provided it will return the smallest argument. If no arguments are provided it will return the smallest element in the stored array"
    Dim localItems() As Variant
    
    If UBound(args) < LBound(args) Then          'Test if no args provided
        localItems = InternalItems
    Else
        localItems = args
    End If
    
    Min = RecursiveMin(localItems)
End Function

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
    Dim Final As Long
    Dim count As Long
    Dim newIndex As Long
    Dim localItems() As Variant
    Dim result() As Variant
    
    localItems = InternalItems
    localLength = this.Length
    relativeStart = startIndex
    If relativeStart < LBound(localItems) Then
        If relativeStart < 0 Then
            oldIndex = Max((localLength + relativeStart), LBound(localItems))
        Else
            oldIndex = Max((localLength - relativeStart), LBound(localItems))
        End If
    Else
        oldIndex = Min(relativeStart, localLength)
    End If
    If endIndex = MISSING_LONG Then
        relativeEnd = localLength + LBound(localItems)
    Else
        relativeEnd = endIndex
    End If
    If relativeEnd < LBound(localItems) Then
        Final = Max((localLength + relativeEnd), LBound(localItems))
    Else
        Final = Min(relativeEnd, localLength + LBound(localItems))
    End If
    
    newIndex = LBound(localItems)
    count = Max(Final - oldIndex, 0) + LBound(localItems)
    If count > newIndex Then
        ReDim result(newIndex To count - 1)
        Do While oldIndex < Final
            If oldIndex >= LBound(localItems) And oldIndex <= UBound(localItems) Then
                LetOrSetElement result(newIndex), localItems(oldIndex)
                newIndex = newIndex + 1
                oldIndex = oldIndex + 1
            End If
        Loop
        If this.ArrayType = BA_MULTIDIMENSION Then
            Slice = JaggedToMulti(result)
        Else
            Slice = result
        End If
    End If
    
End Function

'@Description("Stores the values contained in an Excel range in the internal array")
Public Function FromExcelRange( _
        ByRef fromRange As Object, _
        Optional ByVal detectLastRow As Boolean, _
        Optional ByVal detectLastColumn As Boolean _
    ) As BetterArray
Attribute FromExcelRange.VB_Description = "Stores the values contained in an Excel range in the internal array"

    If TypeName(fromRange) = "Range" Then
        Dim startColumn As Long
        Dim endColumn As Long
        Dim startRow As Long
        Dim endRow As Long
        
        With fromRange
            startColumn = .column
            startRow = .Row
            endColumn = .column + .Columns.count - 1
            endRow = .Row + .Rows.count - 1
        End With
        
        With fromRange.Parent
            If detectLastColumn Then
                endColumn = .Cells.Item(startRow, .Columns.count).End(xlToLeft).column
            End If
            
            If detectLastRow Then
                endRow = .Cells.Item(.Rows.count, startColumn).End(xlUp).Row
            End If
            
            Me.Items = .Range(.Cells(startRow, startColumn), .Cells(endRow, endColumn)).value
        End With
        
        ' convert arrays of single column or row to 1d
        If startColumn = endColumn And startRow <> endRow Then
            Me.Items = Me.ExtractSegment(, startColumn)
        ElseIf startColumn <> endColumn And startRow = endRow Then
            Me.Items = Me.ExtractSegment(startRow)
        End If
        
    Else
        RaiseError ErrorCodes.EC_EXPECTED_RANGE_OBJECT, "FromExcelRange()", "FromRange"
    End If
    
    Set FromExcelRange = Me
End Function

'@Description("Extracts the segment of the array located at the specified indices.")
Public Function ExtractSegment( _
        Optional ByVal rowIndex As Long = MISSING_LONG, _
        Optional ByVal columnIndex As Long = MISSING_LONG _
    ) As Variant()
Attribute ExtractSegment.VB_Description = "Extracts the segment of the array located at the specified indices."
    Dim i As Long
    Dim localRowIndex As Long
    Dim localColumnIndex As Long
    Dim nestedBounds() As Long
    Dim localItems() As Variant
    Dim result() As Variant
    
    localItems = InternalItems

    If rowIndex = MISSING_LONG Then
        If columnIndex = MISSING_LONG Then
            ' if no coords provided return full array
            ' calls to Me.Items instead of returning localItems to covert md from jagged
            result = Me.Items
        Else
            ' Row number missing, column number present
            Select Case this.ArrayType
            Case BA_ONEDIMENSION
                If columnIndex >= LBound(localItems) And columnIndex <= UBound(localItems) Then
                    localColumnIndex = columnIndex
                Else
                    localColumnIndex = LBound(localItems)
                End If
                result = Array(localItems(localColumnIndex))
            Case BA_JAGGED, BA_MULTIDIMENSION
                ' return full column
                nestedBounds = GetMaxBoundsAtDimension(localItems, 2)
                If columnIndex >= nestedBounds(0) And columnIndex <= nestedBounds(1) Then
                    localColumnIndex = columnIndex
                Else
                    localColumnIndex = this.LowerBound
                End If
                ReDim result(LBound(localItems) To UBound(localItems))
                For i = LBound(localItems) To UBound(localItems)
                    result(i) = localItems(i)(localColumnIndex)
                Next
            Case BA_UNALLOCATED
                result = localItems
            Case Else
            End Select
        End If
    Else
        ' Row Number not missing
        If rowIndex >= LBound(localItems) And rowIndex <= UBound(localItems) Then
            localRowIndex = rowIndex
        Else
            localRowIndex = LBound(localItems)
        End If
        If columnIndex = MISSING_LONG Then
            ' Row number present, column number missing
            Select Case this.ArrayType
            Case BA_ONEDIMENSION
                result = Array(localItems(localRowIndex))
            Case BA_JAGGED, BA_MULTIDIMENSION
                result = localItems(localRowIndex)
            Case BA_UNALLOCATED
                result = localItems
            Case Else
            End Select
        Else
            ' Row number present, column number present
            Select Case this.ArrayType
            Case BA_ONEDIMENSION
                ' ignore column number arg as array is 1d
                result = Array(localItems(localRowIndex))
            Case BA_JAGGED, BA_MULTIDIMENSION
                ' return intersection of rowIndex and column number
                nestedBounds = GetMaxBoundsAtDimension(localItems, 2)
                If columnIndex >= nestedBounds(0) And columnIndex <= nestedBounds(1) Then
                    localColumnIndex = columnIndex
                Else
                    localColumnIndex = this.LowerBound
                End If
                If IsArray(localItems(localRowIndex)(localColumnIndex)) Then
                    result = localItems(localRowIndex)(localColumnIndex)
                Else
                    result = Array(localItems(localRowIndex)(localColumnIndex))
                End If
            Case BA_UNALLOCATED
                result = localItems
            Case Else
            End Select
        End If
    End If
    
    ExtractSegment = result
End Function

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
        localItems = InternalItems
        depth = GetJaggedArrayDepth(localItems)
        If depth > 0 Then 'check array is allocated
            If depth = 1 Then
                '1d array
                localItems = EnsureScalar1DArray(localItems)
                If transposeValues Then
                    lengthRows = this.Length
                    localItems = Transpose1DArray(localItems)
                Else
                    lengthColumns = this.Length
                End If
            ElseIf depth > 1 Then
                'jagged array - multidim arrays should also be stored as jagged internally
                Const outputDepth As Long = 2 'need 2 dimensions for tabular representation output
                localItems = JaggedToMulti(localItems, outputDepth, ensureScalar:=True)
                If transposeValues Then
                    lengthRows = UBound(localItems, 2) - LBound(localItems, 2) + 1
                    lengthColumns = UBound(localItems, 1) - LBound(localItems, 1) + 1
                    localItems = Transpose2DArray(localItems)
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
        '@Ignore EmptyElseBlock
        Else
            'array must be unalloc
            'do nothing(?)
        End If
    Else
        RaiseError ErrorCodes.EC_EXPECTED_RANGE_OBJECT, "ToExcelRange()", "Destination"
    End If
    
    Set ToExcelRange = localRange
End Function

'@Description("Tests if the stored array is sorted in ascending order.")
Public Function IsSorted(Optional ByVal columnIndex As Long = MISSING_LONG) As Boolean
Attribute IsSorted.VB_Description = "Tests if the stored array is sorted in ascending order."
    Dim i As Long
    Dim localLowerBound As Long
    Dim localUpperBound As Long
    Dim localColumnIndex As Long
    Dim depth As Long
    Dim result As Boolean
    Dim localItems() As Variant
    Dim storedType As ArrayTypes
    
    result = True
    
    localItems = InternalItems
    localLowerBound = LBound(localItems)
    localUpperBound = UBound(localItems)
    storedType = GetArrayType(localItems)
    If storedType <> BA_UNDEFINED And storedType <> BA_UNALLOCATED Then
        Select Case storedType
        Case BA_ONEDIMENSION
            For i = localLowerBound To localUpperBound - 1
                If localItems(i) > localItems(i + 1) Then
                    result = False
                    Exit For
                End If
            Next
        Case BA_MULTIDIMENSION
            ' should never be true as all md arrays stored as jagged internally
            If columnIndex = MISSING_LONG Then
                localColumnIndex = LBound(localItems, 2)
            Else
                localColumnIndex = CLng(columnIndex)
            End If
            For i = localLowerBound To localUpperBound - 1
                If localItems(i, localColumnIndex) > localItems(i + 1, localColumnIndex) Then
                    result = False
                    Exit For
                End If
            Next
        Case BA_JAGGED
            depth = GetJaggedArrayDepth(localItems)
            If depth > 2 Then
                IsSorted = False
                RaiseError EC_EXCEEDS_MAX_SORT_DEPTH, "IsSorted"
            Else
                For i = localLowerBound To localUpperBound - 1
                    If i = localLowerBound Then
                        If columnIndex = MISSING_LONG Then
                            localColumnIndex = LBound(localItems(i))
                        Else
                            localColumnIndex = CLng(columnIndex)
                        End If
                    End If
                    If localItems(i)(localColumnIndex) > localItems(i + 1)(localColumnIndex) Then
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
        Optional ByVal fromIndex As Long = MISSING_LONG, _
        Optional ByVal compType As ComparisonType _
    ) As Long
Attribute IndexOf.VB_Description = "returns the first index at which a given element can be found in the array"
    Dim localItems() As Variant
    Dim relativeStart As Long
    Dim currentIndex As Long
    
    If compType = CT_LIKENESS Then
        If TypeName(searchElement) <> "String" Then
            RaiseError EC_STRING_TYPE_EXPECTED, "IndexOf", "searchElement"
        End If
    End If
    
    If this.Length = 0 Then
        IndexOf = MISSING_LONG
        Exit Function
    End If

    localItems = InternalItems

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
    
    Dim isMatch As Boolean
    Do While currentIndex <= UBound(localItems)
        Select Case compType
        Case ComparisonType.CT_LIKENESS
            isMatch = CStr(localItems(currentIndex)) Like CStr(searchElement)
        Case Else
            isMatch = ElementsAreEqual(searchElement, localItems(currentIndex))
        End Select
        If isMatch Then
            IndexOf = currentIndex
            Exit Function
        End If
        currentIndex = currentIndex + 1
    Loop
    
    IndexOf = MISSING_LONG
End Function

'@Description("returns the first index at which a given element can be found in the array")
Public Function LastIndexOf( _
        ByVal searchElement As Variant, _
        Optional ByVal fromIndex As Long = MISSING_LONG, _
        Optional ByVal compType As ComparisonType _
    ) As Long
Attribute LastIndexOf.VB_Description = "returns the first index at which a given element can be found in the array"
    Dim localItems() As Variant
    Dim currentIndex As Long
    
    If compType = CT_LIKENESS Then
        If TypeName(searchElement) <> "String" Then
            RaiseError EC_STRING_TYPE_EXPECTED, "IndexOf", "searchElement"
        End If
    End If
    
    If this.Length = 0 Then
        LastIndexOf = MISSING_LONG
        Exit Function
    End If

    localItems = InternalItems
    
    If fromIndex = MISSING_LONG Then
        currentIndex = UBound(localItems)
    Else
        If fromIndex >= LBound(localItems) Then
            currentIndex = Min(fromIndex, UBound(localItems))
        ElseIf fromIndex < 0 Then
            currentIndex = UBound(localItems) + fromIndex
        End If
        
    End If
    
    Dim isMatch As Boolean
    Do While currentIndex >= LBound(localItems)
        Select Case compType
        Case ComparisonType.CT_LIKENESS
            isMatch = CStr(localItems(currentIndex)) Like CStr(searchElement)
        Case Else
            isMatch = ElementsAreEqual(searchElement, localItems(currentIndex))
        End Select
        If isMatch Then
            LastIndexOf = currentIndex
            Exit Function
        End If
        currentIndex = currentIndex - 1
    Loop
    
    LastIndexOf = MISSING_LONG
End Function

'@Description("Removes the item at the specified index. Returns the new length of the array")
Public Function Remove(ByVal index As Long) As Long
Attribute Remove.VB_Description = "Removes the item at the specified index. Returns the new length of the array"
    Dim relativeIndex As Long
    Dim localType As ArrayTypes
    
    relativeIndex = MISSING_LONG
    localType = this.ArrayType
    
    If index >= this.LowerBound Then
        If index <= Me.UpperBound Then
            relativeIndex = index
        End If
    Else
        If index < 0 Then
            relativeIndex = Me.UpperBound + index
            If relativeIndex < this.LowerBound Then relativeIndex = MISSING_LONG
        End If
    End If
    
    If relativeIndex <> MISSING_LONG Then
        Dim beforeSlice() As Variant
        Dim afterSlice() As Variant
        Dim beforeExists As Boolean
        Dim afterExists As Boolean
        
        If relativeIndex > this.LowerBound Then
            beforeSlice = Me.Slice(this.LowerBound, relativeIndex)
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
            If this.ArrayType = BA_JAGGED And localType = BA_MULTIDIMENSION Then
                this.ArrayType = localType
            End If
        End If
    End If
    
    Remove = this.Length
End Function

'@Description("changes the contents of the array by removing or replacing existing elements and/or adding new elements in place.")
Public Function Splice( _
        ByVal startIndex As Long, _
        ParamArray args() As Variant _
    ) As Variant()
Attribute Splice.VB_Description = "changes the contents of the array by removing or replacing existing elements and/or adding new elements in place."
    Dim localItems() As Variant
    Dim actualStart As Long
    Dim actualDeleteCount As Long
    Dim argsCount As Long
    Dim localLength As Long
    Dim localArgs() As Variant
    Dim result() As Variant
    Dim i As Long
    Dim tempArray() As Variant
    Dim itemCount As Long
    Dim tempItems() As Variant
    
    localArgs = args
    localItems = InternalItems
    localLength = UBound(localItems) + 1
    If startIndex < LBound(localItems) Then
        actualStart = Max(localLength + startIndex, LBound(localItems))
    Else
        actualStart = Min(startIndex, localLength)
    End If
    argsCount = GetArrayLength(localArgs)
    If argsCount = 0 Then
        ' only a start index has been supplied
        ' we're gonna delete everything
        itemCount = 0
        actualDeleteCount = localLength - actualStart
    Else
        ' first arg in args is delete count, second onwards are insertion items
        itemCount = argsCount - 1
        'deleteCount is first element in args.
        actualDeleteCount = Min(Max(args(LBound(args)), 0), localLength - actualStart)
    End If
    
    If localLength + itemCount - actualDeleteCount > MAX_ARRAY_LENGTH Then
        RaiseError EC_EXCEEDS_MAX_ARRAY_LENGTH, "Splice"
        Exit Function
    End If

    ' get items to be deleted for return val
    If actualDeleteCount > 0 Then
        ReDim result(0 To actualDeleteCount - 1)
        i = 0
        Do While i < actualDeleteCount
            LetOrSetElement result(i), localItems(actualStart + i)
            i = i + 1
        Loop
    Else
        ' return empty array if nothing to be deleted
        ReDim result(0)
    End If
    
    If itemCount > 0 Then
        ReDim tempItems(0 To itemCount - 1)
        For i = LBound(tempItems) To UBound(tempItems)
            tempItems(i) = args(i + 1)
        Next
    End If
            
    If itemCount < actualDeleteCount Then
        i = actualStart
        Do While i < (localLength - actualDeleteCount)
            LetOrSetElement localItems(i + itemCount), localItems(i + actualDeleteCount)
            i = i + 1
        Loop
        
        
        i = localLength
        Dim tempBetterArray As BetterArray
        Set tempBetterArray = New BetterArray
        tempBetterArray.Items = localItems
        Do While i > (localLength - actualDeleteCount + itemCount)
            tempBetterArray.Remove (i - 1)
            ' Perform ? DeletePropertyOrThrow(O, ! ToString(k - 1)).
            i = i - 1
        Loop
        localItems = tempBetterArray.Items
        Set tempBetterArray = Nothing
    ElseIf itemCount > actualDeleteCount Then
        i = (localLength - actualDeleteCount)
        tempArray = localItems
        ReDim Preserve tempArray(this.LowerBound To i + itemCount - 1) As Variant
        Do While i > actualStart
            LetOrSetElement tempArray(i + itemCount - 1), localItems(i + actualDeleteCount - 1)
            i = i - 1
        Loop
        localItems = tempArray
    End If
    
    If itemCount > 0 Then
        i = actualStart
        Dim j As Long
        For j = LBound(tempItems) To UBound(tempItems)
            localItems(i) = tempItems(j)
            i = i + 1
        Next
    End If
    
    InternalItems = localItems
    
    Splice = result
End Function


''''''''''''''''''''''''''''
' Chainable Public Methods '
''''''''''''''''''''''''''''

'@Description("Recursvely fills (modifies) all the elements of an array from a start index (default zero) to an end index (default array length) with a passed value.")
Public Function Fill( _
        ByVal value As Variant, _
        Optional ByVal startIndex As Long = MISSING_LONG, _
        Optional ByVal endIndex As Long = MISSING_LONG _
    ) As BetterArray
Attribute Fill.VB_Description = "Recursvely fills (modifies) all the elements of an array from a start index (default zero) to an end index (default array length) with a passed value."
    Dim localItems() As Variant
    Dim relativeStart As Long
    Dim relativeEnd As Long
    
    localItems = InternalItems
    If startIndex = MISSING_LONG Then
        relativeStart = LBound(localItems)
    Else
        If startIndex < 0 Then
            relativeStart = Max(Me.UpperBound + startIndex, LBound(localItems))
        ElseIf startIndex < LBound(localItems) Then
            relativeStart = LBound(localItems)
        Else
            relativeStart = Min(startIndex, Me.UpperBound)
        End If
    End If
    
    If endIndex = MISSING_LONG Then
        relativeEnd = Me.UpperBound
    Else
        If endIndex < 0 And endIndex < LBound(localItems) Then
            relativeEnd = Max(Me.UpperBound + endIndex, LBound(localItems))
        ElseIf endIndex < LBound(localItems) Then
            relativeEnd = Me.UpperBound
        Else
            relativeEnd = Min(endIndex, Me.UpperBound)
        End If
    End If
    InternalItems = RecursiveFill(localItems, value, relativeStart, relativeEnd)
    Set Fill = Me
End Function

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

'@Description("Swaps rows for columns & vice versa.")
Public Function Transpose() As BetterArray
Attribute Transpose.VB_Description = "Swaps rows for columns & vice versa."
    Dim result() As Variant
    Dim localItems() As Variant
    Dim localType As ArrayTypes
    Dim storedType As ArrayTypes
    storedType = this.ArrayType
    localItems = InternalItems
    localType = GetArrayType(localItems)
    Select Case localType
    Case ArrayTypes.BA_ONEDIMENSION
        result = Transpose1DArray(localItems)
    Case ArrayTypes.BA_MULTIDIMENSION
        ' Sould Never match as MD arrs stored as jagged internally
        result = Transpose2DArray(localItems)
    Case ArrayTypes.BA_JAGGED
        result = TransposeArrayOfArrays(localItems)
    Case Else
        result = localItems
    End Select
    Me.Items = result
    If storedType = BA_MULTIDIMENSION Then
        localType = GetArrayType(result)
        If localType = BA_JAGGED Then
            this.ArrayType = BA_MULTIDIMENSION
        End If
    End If
    Set Transpose = Me
End Function

'@Description("Returns a new BetterArray instance containing the same values as the current instance.")
Public Function Clone() As BetterArray
Attribute Clone.VB_Description = "Returns a new BetterArray instance containing the same values as the current instance."
    Dim result As BetterArray
    Set result = New BetterArray
    result.LowerBound = this.LowerBound
    result.Items = Me.Items
    Set Clone = result
End Function

'@Description("Clears the current array and resets capacity to default value")
Public Function ResetToDefault() As BetterArray
Attribute ResetToDefault.VB_Description = "Clears the current array and resets capacity to default value"
    this.LowerBound = 0
    ReDim this.Items(this.LowerBound To DEFAULT_CAPACITY + this.LowerBound)
    this.Length = 0
    Me.Capacity = DEFAULT_CAPACITY
    this.ArrayType = BA_UNALLOCATED
    Set ResetToDefault = Me
End Function

'@Description("Clears all entries in the current array but retains existing capacity.")
Public Function Clear() As BetterArray
Attribute Clear.VB_Description = "Clears all entries in the current array but retains existing capacity."
    Dim oldCapacity As Long
    oldCapacity = this.Capacity
    ReDim this.Items(this.LowerBound To oldCapacity + this.LowerBound)
    this.ArrayType = BA_UNALLOCATED
    this.Length = 0
    Me.Capacity = oldCapacity
    Set Clear = Me
End Function

'@Description("Joins one or more arrays onto the end of the current array.")
Public Function Concat(ParamArray args() As Variant) As BetterArray
Attribute Concat.VB_Description = "Joins one or more arrays onto the end of the current array."
    Dim arg As Variant
    Dim localItems() As Variant
    Dim localType As ArrayTypes
    Dim i As Long
    
    For Each arg In args
        If IsArray(arg) Then
            localItems = arg
            ' should never return BA_UNDEFINED as we've already checked it's an array
            localType = GetArrayType(localItems)
            ' If pushing to unallocated array short circuit individual pushes and just assign with me.items
            If this.ArrayType = BA_UNALLOCATED Then
                If localType = BA_UNALLOCATED Then
                    Me.Items = GetEmptyArray
                Else
                    Me.Items = localItems
                End If
            Else
                If localType = BA_MULTIDIMENSION Then
                    localItems = MultiToJagged(localItems)
                End If
                
                If this.ArrayType = BA_ONEDIMENSION Then
                    ' if the stored array is not already a jagged array and
                    ' the passed array is a MD array then set our array type to MD
                    If localType = BA_MULTIDIMENSION And this.ArrayType <> BA_JAGGED Then
                        this.ArrayType = BA_MULTIDIMENSION
                    End If
                End If
                
                For i = LBound(localItems) To UBound(localItems)
                    Me.Push localItems(i)
                Next
            End If
        Else
            Me.Push arg
        End If
    Next
    Set Concat = Me
End Function

'@Description("Writes all of the elements in a Collection objet to the internal array.")
Public Function CopyFromCollection(ByVal sourceCollection As Collection) As BetterArray
Attribute CopyFromCollection.VB_Description = "Writes all of the elements in a Collection objet to the internal array."
    If sourceCollection Is Nothing Then
        RaiseError EC_EXPECTED_COLLECTION_OBJECT, "CopyFromCollection", "SourceCollection"
    End If
    Dim i As Long
    Dim newItems() As Variant
    this.Length = sourceCollection.count
    If this.Length = 0 Then
        newItems = GetEmptyArray
    Else
        ReDim newItems(this.LowerBound To (this.Length - this.LowerBound - 1))
        For i = 1 To this.Length
            newItems(i + this.LowerBound - 1) = sourceCollection.Item(i)
        Next
    End If
    Me.Items = newItems
    Set CopyFromCollection = Me
End Function

'@Description("Sorts the stored array in ascending order.")
Public Function Sort(Optional ByVal sortColumn As Long = MISSING_LONG) As BetterArray
Attribute Sort.VB_Description = "Sorts the stored array in ascending order."
    Dim localItems() As Variant
    Dim sortedItems() As Variant
    Dim localArrayType As ArrayTypes
    
    localItems = InternalItems
    localArrayType = GetArrayType(localItems)
    If localArrayType = ArrayTypes.BA_UNALLOCATED Or _
       localArrayType = ArrayTypes.BA_UNDEFINED Then
        sortedItems = GetEmptyArray
    Else
        If this.Length > 0 Then
            If localArrayType <> BA_ONEDIMENSION Then
                If localArrayType = ArrayTypes.BA_MULTIDIMENSION Then
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
                localArrayType, _
                IIf(sortColumn = MISSING_LONG, LBound(localItems), sortColumn)
        End If
        sortedItems = localItems
    End If
    Me.Items = sortedItems
    Set Sort = Me
End Function

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
    Dim Final As Long
    Dim count As Long
    Dim direction As Long
    Dim localItems() As Variant
    
    Select Case this.ArrayType
    Case ArrayTypes.BA_UNDEFINED
        RaiseError EC_EXPECTED_ARRAY, "CopyWithin"
    Case ArrayTypes.BA_UNALLOCATED
        RaiseError EC_UNALLOCATED_ARRAY, "CopyWithin"
    Case Else
        localItems = InternalItems
        localLength = this.Length
        relativeTarget = target
        
        If relativeTarget < 0 Then
            toIndex = Max((localLength + relativeTarget), 0)
        Else
            toIndex = Min(relativeTarget, localLength)
        End If
        
        relativeStart = IIf(startIndex = MISSING_LONG, LBound(localItems), startIndex)
        
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
            Final = Max((localLength + relativeEnd), 0)
        Else
            Final = Min(relativeEnd, localLength)
        End If
        
        count = Min(Final - fromIndex, localLength - toIndex)
        
        If fromIndex < toIndex And toIndex < fromIndex + count Then
            direction = -1
            fromIndex = fromIndex + count - 1
            toIndex = toIndex + count - 1
        Else
            direction = 1
        End If
        
        Do While count > 0
            If fromIndex >= LBound(localItems) And fromIndex <= UBound(localItems) Then
                localItems(toIndex) = localItems(fromIndex)
            End If
            fromIndex = fromIndex + direction
            toIndex = toIndex + direction
            count = count - 1
        Loop
        
        Me.Items = localItems
    End Select
    
    Set CopyWithin = Me

End Function

'@Ignore ShadowedDeclaration
'@Description("Filters the stored array to include only values that meet the specified criteria.")
Public Function Filter( _
        ByVal match As Variant, _
        Optional ByVal include As Boolean, _
        Optional ByVal recurse As Boolean) As BetterArray
Attribute Filter.VB_Description = "Filters the stored array to include only values that meet the specified criteria."
    Dim result() As Variant
    Dim localItems() As Variant
    Dim localType As ArrayTypes
    localType = this.ArrayType
    If localType = BA_UNDEFINED Then
        RaiseError EC_EXPECTED_ARRAY, "Filter"
    ElseIf localType = BA_UNALLOCATED Then
        result = GetEmptyArray
    Else
        localItems = InternalItems
        result = RecursiveFilter( _
            localItems, _
            match, _
            include, _
            recurse _
        )
    End If
    Me.Items = result
    If localType = BA_MULTIDIMENSION Then this.ArrayType = localType
    Set Filter = Me
End Function


'@Description("Filters the stored array to include only values of the specified type.")
Public Function FilterType( _
        ByVal searchTypeName As String, _
        Optional ByVal include As Boolean, _
        Optional ByVal recurse As Boolean) As BetterArray
Attribute FilterType.VB_Description = "Filters the stored array to include only values of the specified type."
    Dim result() As Variant
    Dim localItems() As Variant
    Dim localType As ArrayTypes
    localType = this.ArrayType
    If localType = BA_UNDEFINED Then
        RaiseError EC_EXPECTED_ARRAY, "Filter"
    ElseIf localType = BA_UNALLOCATED Then
        result = GetEmptyArray
    Else
        localItems = InternalItems
        result = RecursiveFilter( _
            localItems, _
            searchTypeName, _
            include, _
            recurse, _
            True _
        )
    End If
    Me.Items = result
    If localType = BA_MULTIDIMENSION Then this.ArrayType = localType
    Set FilterType = Me
End Function

'@Description("Reverses the order of elements in the array. The first array element becomes the last, and the last array element becomes the first.")
Public Function Reverse(Optional ByVal recurse As Boolean) As BetterArray
Attribute Reverse.VB_Description = "Reverses the order of elements in the array. The first array element becomes the last, and the last array element becomes the first."
    Dim localItems() As Variant
    Dim result() As Variant
    Dim currentType As ArrayTypes
    currentType = this.ArrayType
    Select Case currentType
    Case BA_UNDEFINED
        RaiseError EC_UNDEFINED_ARRAY, "Reverse"
    Case BA_ONEDIMENSION, BA_MULTIDIMENSION, BA_JAGGED
        localItems = InternalItems
        result = RecursiveReverse(localItems, recurse)
        Me.Items = result
        this.ArrayType = currentType
    End Select
    
    Set Reverse = Me
End Function

'@Description("Flattens a Multi-Dimensional or Jagged array into One Dimension arrays.")
Public Function Flatten() As BetterArray
Attribute Flatten.VB_Description = "Flattens a Multi-Dimensional or Jagged array into One Dimension arrays."
    Dim localItems() As Variant
    Dim result As BetterArray
        
    ' noop on 1d or unalloc
    ' md & jagged both stored as jagged
    Select Case this.ArrayType
    Case BA_UNDEFINED
        RaiseError EC_UNDEFINED_ARRAY, "Flatten"
    Case BA_MULTIDIMENSION, BA_JAGGED
        localItems = InternalItems
        Set result = New BetterArray
        result.LowerBound = this.LowerBound
        RecursiveFlatten localItems, result
        Me.Items = result.Items
    End Select

    Set Flatten = Me
End Function

'@Description("shuffles the order of the stored array.")
Public Function Shuffle(Optional ByVal recurse As Boolean) As BetterArray
Attribute Shuffle.VB_Description = "shuffles the order of the stored array."
    Dim localItems() As Variant
    Dim result() As Variant
    Dim currentType As ArrayTypes
    currentType = this.ArrayType
    Select Case currentType
    Case BA_UNDEFINED
        RaiseError EC_UNDEFINED_ARRAY, "Shuffle"
    Case BA_ONEDIMENSION, BA_MULTIDIMENSION, BA_JAGGED
        localItems = InternalItems
        result = RecursiveShuffle(localItems, recurse)
        Me.Items = result
        this.ArrayType = currentType
    End Select
    
    Set Shuffle = Me
End Function

'@Description("filters the array to contain only unique elements")
Public Function Unique(Optional ByVal columnIndex As Long = MISSING_LONG) As BetterArray
Attribute Unique.VB_Description = "filters the array to contain only unique elements"
    Dim localItems() As Variant
    Dim localType As ArrayTypes
    Dim result As BetterArray
    Dim i As Long
    
    localItems = InternalItems
    localType = this.ArrayType
    Set result = New BetterArray
    If (localType = BA_JAGGED Or localType = BA_MULTIDIMENSION) And columnIndex <> MISSING_LONG Then
        Dim localDepth As Long
        localDepth = GetJaggedArrayDepth(localItems)
        If localDepth = 2 Then
            ' if array is 2d we can compare unique based on specified column
            Dim localColumn As Long
            Dim comparisonValues() As Variant
            Dim comparator As BetterArray
            Set comparator = New BetterArray
            Dim bounds() As Long
            
            bounds = GetMaxBoundsAtDimension(localItems, 2)
            
            ' determine column to compare - columnIndex if valid, lbound if not
            If columnIndex > 0 And columnIndex <= (bounds(1) - bounds(0) + 1) Then
                localColumn = columnIndex - 1 + bounds(0)
            Else
                localColumn = bounds(0)
            End If
            
            ' extract the column
            comparisonValues = Me.ExtractSegment(columnIndex:=localColumn)
            ' store the extracted column in a new BetterArray instance
            comparator.Items = comparisonValues
            
            For i = LBound(localItems) To UBound(localItems)
                If comparator.IndexOf(localItems(i)(localColumn)) = i Then
                    result.Push localItems(i)
                End If
            Next
            
        Else
            ' array has more then two dimensions
            ' just return unique in outermost array
            For i = LBound(localItems) To UBound(localItems)
                If Me.IndexOf(localItems(i)) = i Then
                    result.Push localItems(i)
                End If
            Next
        
        End If
    Else
        ' no columnIndex arg or array is one-dimension
        ' unique in outermost array
        For i = LBound(localItems) To UBound(localItems)
            If Me.IndexOf(localItems(i)) = i Then
                result.Push localItems(i)
            End If
        Next
    End If
    

    Me.Items = result.Items
    this.ArrayType = localType
    Set Unique = Me
End Function

'@Description("Accepts a path argument pointing to a comma-separated values (CSV) file and stores the delimited values contained within to the internal array.")
Public Function FromCSVFile( _
    ByVal path As String, _
    Optional ByVal ColumnDelimiter As String = ",", _
    Optional ByVal RowDelimiter As String = vbCrLf, _
    Optional ByVal Quote As String = """", _
    Optional ByVal IgnoreFirstRow As Boolean, _
    Optional ByVal DuckType As Boolean _
) As BetterArray
Attribute FromCSVFile.VB_Description = "Accepts a path argument pointing to a comma-separated values (CSV) file and stores the delimited values contained within to the internal array."
    Dim rawData As String
       
    Open path For Input As #1
        rawData = Input$(LOF(1), #1)
    Close #1
       
    Set FromCSVFile = FromCSVString( _
        CSVString:=rawData, _
        ColumnDelimiter:=ColumnDelimiter, _
        RowDelimiter:=RowDelimiter, _
        Quote:=Quote, _
        IgnoreFirstRow:=IgnoreFirstRow, _
        DuckType:=DuckType _
    )
End Function

'@Description("Accepts a string argument representing the contents of a comma-separated values (CSV) file and stores the delimited values contained within to the internal array.")
Public Function FromCSVString( _
    ByVal CSVString As String, _
    Optional ByVal ColumnDelimiter As String = ",", _
    Optional ByVal RowDelimiter As String = vbCrLf, _
    Optional ByVal Quote As String = """", _
    Optional ByVal IgnoreFirstRow As Boolean, _
    Optional ByVal DuckType As Boolean _
) As BetterArray
Attribute FromCSVString.VB_Description = "Accepts a string argument representing the contents of a comma-separated values (CSV) file and stores the delimited values contained within to the internal array."
    Dim localType As ArrayTypes
    Dim parsed() As Variant
    
    If this.ArrayType = BA_MULTIDIMENSION Then
        localType = BA_MULTIDIMENSION
    Else
        localType = BA_JAGGED
    End If
     
    parsed = ParseCSV( _
        Expression:=CSVString, _
        ColumnDelimiter:=ColumnDelimiter, _
        RowDelimiter:=RowDelimiter, _
        Quote:=Quote, _
        IgnoreFirstRow:=IgnoreFirstRow, _
        ReturnJagged:=IIf(localType = BA_JAGGED, True, False), _
        Base:=this.LowerBound, _
        DuckType:=DuckType _
    )

    this.Length = GetArrayLength(parsed)
    this.Capacity = UBound(parsed)
    this.ArrayType = localType
    this.Items = parsed

    Set FromCSVString = Me
End Function

'''''''''''''''''''
' Private Methods '
'''''''''''''''''''

'@Description("Populates TString Types")
Private Function StringFactory(ByRef Expression As String) As TString
Attribute StringFactory.VB_Description = "Populates TString Types"
    Dim result As TString
    result.Text = Expression
    result.Length = Len(Expression)
    result.ByteLength = LenB(Expression)
    StringFactory = result
End Function

'@Description("Returns the next byte position in the string of the specified delimiter")
Private Function NextDelimBytePos(ByRef Expression As TString, ByRef Delimiter As TString, Optional ByRef startIndex As Long = 1) As Long
Attribute NextDelimBytePos.VB_Description = "Returns the next byte position in the string of the specified delimiter"
    Dim result As Long
    result = InStrB(startIndex, Expression.Text, Delimiter.Text, vbBinaryCompare)
    Do Until (result And 1) Or (result = 0)
        result = InStrB(result + 1, Expression.Text, Delimiter.Text, vbBinaryCompare)
    Loop
    NextDelimBytePos = result
End Function

'@Description("Splits the provided string into rows based on the provided row delimiter")
Private Function GetCSVRows( _
    ByRef Expr As TString, _
    ByRef RowDelim As TString, _
    ByRef LiteralDelim As TString _
) As String()
Attribute GetCSVRows.VB_Description = "Splits the provided string into rows based on the provided row delimiter"
    Dim rowDelimIndex As Long
    Dim lastRowDelim As Long
    Dim rowIndex As Long
    Dim blankTrailingRows As Long
    Dim maxRows As Long
    Dim lineEndIndices() As Long
    Dim literalDelimIndex As Long
    Dim countRows As Long
    Dim i As Long
    Dim cursor As Long
    Dim CSVRows() As String

    ' find first row delimiter
    rowDelimIndex = NextDelimBytePos(Expr, RowDelim)

    ' find last row delimiter
    lastRowDelim = InStrRev(Expr.Text, RowDelim.Text)
    rowIndex = Expr.Length + 1
    Do While rowIndex - lastRowDelim = RowDelim.Length
        rowIndex = lastRowDelim
        lastRowDelim = InStrRev(Expr.Text, RowDelim.Text, lastRowDelim)
        blankTrailingRows = blankTrailingRows + 1
    Loop
    rowIndex = Empty

    ' estimate max number of rows - does not account for line breaks within strings
    maxRows = (Expr.Length - Len(Replace(Expr.Text, RowDelim.Text, vbNullString))) / RowDelim.Length + 1
    ReDim lineEndIndices(0 To maxRows)
        
    ' find first quote delimiter
    literalDelimIndex = NextDelimBytePos(Expr, LiteralDelim)
    
    Do While rowDelimIndex > 0
        ' see if row delimiter before quote delimiter
        If rowDelimIndex + RowDelim.ByteLength <= literalDelimIndex Or literalDelimIndex = 0 Then
            ' no preceding literalDelim - iterate row count
            lineEndIndices(countRows) = rowDelimIndex
            rowDelimIndex = NextDelimBytePos(Expr, RowDelim, rowDelimIndex + RowDelim.ByteLength)
            countRows = countRows + 1
        Else
            ' quote delimiter before row delimiter
            ' find next quote delim
            literalDelimIndex = NextDelimBytePos(Expr, LiteralDelim, literalDelimIndex + 2)
           ' see if closing quote delim found
            If literalDelimIndex Then
                ' find next row delim
                rowDelimIndex = NextDelimBytePos(Expr, RowDelim, literalDelimIndex + 2)
                ' if next row delim found
                If rowDelimIndex Then
                    ' find next quote delim
                    literalDelimIndex = NextDelimBytePos(Expr, LiteralDelim, literalDelimIndex + 2)
                End If
            End If
        End If
    Loop
    
    lineEndIndices(countRows) = Expr.ByteLength + 1
    countRows = countRows - blankTrailingRows + 1
    ReDim CSVRows(0 To countRows - 1)
    cursor = 1
    For i = 0 To countRows - 1
        CSVRows(i) = MidB$(Expr.Text, cursor, lineEndIndices(i) - cursor)
        cursor = lineEndIndices(i) + RowDelim.ByteLength
    Next
    GetCSVRows = CSVRows
End Function

'@Description("Counts the number of delimited columns in a provided string")
Private Function CountCSVColumns(ByRef FirstRow As String, ByRef LiteralDelim As TString, ByVal ColumnDelimiter As String) As Long
Attribute CountCSVColumns.VB_Description = "Counts the number of delimited columns in a provided string"
    ' Each line should contain the same number of fields
    ' https://www.loc.gov/preservation/digital/formats/fdd/fdd000323.shtml
    ' so we can just count the columnd in the first row
    Dim i As Long
    Dim inQuote As Boolean
    Dim countColumns As Long
    Dim char As String
    countColumns = 1
    For i = 0 To Len(FirstRow)
        char = Mid$(FirstRow, i + 1, 1)
        Select Case char
            Case LiteralDelim.Text
                inQuote = Not inQuote
            Case ColumnDelimiter
                If Not inQuote Then
                    countColumns = countColumns + 1
                End If
        End Select
    Next
    CountCSVColumns = countColumns
End Function

'@Description("Parses CSV string into a 2d or jagged array")
Private Function ParseCSV( _
    ByRef Expression As String, _
    ByVal ColumnDelimiter As String, _
    ByVal RowDelimiter As String, _
    ByRef Quote As String, _
    ByVal IgnoreFirstRow As Boolean, _
    ByVal ReturnJagged As Boolean, _
    ByVal Base As Long, _
    ByVal DuckType As Boolean _
) As Variant()
Attribute ParseCSV.VB_Description = "Parses CSV string into a 2d or jagged array"
    ' CSV is a simple format for representing a rectangular array (matrix) of numeric and textual values.
    ' It an example of a "flat file" format.
    ' It is a delimited data format that has fields/columns separated by the comma character %x2C (Hex 2C)
    ' and records/rows/lines separated by characters indicating a line break.
    ' RFC 4180 stipulates the use of CRLF pairs to denote line breaks, where CR is %x0D (Hex 0D) and LF is %x0A (Hex 0A).
    ' Each line should contain the same number of fields.
    ' Fields that contain a special character (comma, CR, LF, or double quote), must be "escaped" by enclosing them in double quotes (Hex 22).
    ' An optional header line may appear as the first line of the file with the same format as normal record lines.
    ' This header will contain names corresponding to the fields in the file and should contain the same number of fields as the records in the rest of the file.
    ' CSV commonly employs US-ASCII as character set, but other character sets are permitted.
    ' https://www.loc.gov/preservation/digital/formats/fdd/fdd000323.shtml

    Dim i As Long
    Dim j As Long
    Dim cursor As Long
    Dim columnIndex As Long
    Dim countColumns As Long
    Dim countRows As Long
    Dim lastRow As Long
    Dim lastColumn As Long
    Dim startRow As Long
    Dim inQuote As Boolean
    Dim buffer As String
    Dim char As String
    Dim frag As String
    Dim CSVRows() As String
    Dim jaggedRow() As Variant
    Dim results() As Variant
    Dim element As Variant
    Dim Expr As TString
    Dim RowDelim As TString
    Dim LiteralDelim As TString
    
    Expr = StringFactory(Expression)
    RowDelim = StringFactory(RowDelimiter)
    LiteralDelim = StringFactory(Quote)
    
    If Expr.ByteLength = 0 Or RowDelim.ByteLength = 0 Then
        ' TODO: Raise Err?
        ParseCSV = Array()
        Exit Function
    End If

    CSVRows = GetCSVRows(Expr, RowDelim, LiteralDelim)
        
    countColumns = CountCSVColumns(CSVRows(0), LiteralDelim, ColumnDelimiter)
    countRows = UBound(CSVRows) - LBound(CSVRows) + 1
    
    If IgnoreFirstRow Then
        startRow = 1
    Else
        startRow = 0
    End If
    
    lastRow = Base + countRows - 1 - startRow
    lastColumn = Base + countColumns - 1
    If ReturnJagged Then
        ReDim jaggedRow(Base To lastColumn)
        ReDim results(Base To lastRow)
    Else
        ReDim results(Base To lastRow, Base To lastColumn)
    End If
    
    For i = startRow To lastRow
        columnIndex = Base
        buffer = CSVRows(i)
        cursor = 1
        For j = 0 To Len(buffer)
            char = Mid$(buffer, j + 1, 1)
            If char = LiteralDelim.Text Then
                inQuote = Not inQuote
            ElseIf char = ColumnDelimiter Or j = Len(buffer) Then
                If Not inQuote Then
                    frag = Mid$(buffer, cursor, j - cursor + 1)
                    If DuckType Then
                        element = DuckTypeElement(frag)
                    Else
                        element = UnquoteString(frag)
                    End If
                    If ReturnJagged Then
                        jaggedRow(columnIndex) = element
                    Else
                        results(i - startRow, columnIndex) = element
                    End If
                    j = j + 1
                    cursor = j + 1
                    columnIndex = columnIndex + 1
                    frag = vbNullString
                End If
            End If
        Next
        If ReturnJagged Then
            results(i - startRow) = jaggedRow
            ReDim jaggedRow(Base To lastColumn)
        End If
    Next
    
    Debug.Print Format$(lastRow, "#,###") & " rows parsed"
    ParseCSV = results
End Function

'@Description("Recursvely fills the array (and nested arrays) with the passed value.")
Private Function RecursiveFill( _
        ByRef sourceArray() As Variant, _
        ByVal value As Variant, _
        Optional ByVal startIndex As Long = MISSING_LONG, _
        Optional ByVal endIndex As Long = MISSING_LONG _
    ) As Variant()
Attribute RecursiveFill.VB_Description = "Recursvely fills the array (and nested arrays) with the passed value."
    Dim relativeStart As Long
    Dim relativeEnd As Long
    Dim i As Long
    
    If startIndex = MISSING_LONG Then
        relativeStart = LBound(sourceArray)
    Else
        relativeStart = startIndex
    End If
    If endIndex = MISSING_LONG Then
        relativeEnd = UBound(sourceArray)
    Else
        relativeEnd = endIndex
    End If
    For i = relativeStart To relativeEnd
        If IsArray(sourceArray(i)) Then
            Dim passThru() As Variant
            passThru = RecursiveFill(passThru, value)
        Else
            LetOrSetElement sourceArray(i), value
        End If
    Next
    RecursiveFill = sourceArray
End Function

'@Description("Recursively checks if the passed array only includes the search element.")
Private Function RecursiveEvery( _
        ByVal searchElement As Variant, _
        ByRef searchArray() As Variant, _
        Optional ByVal fromIndex As Long, _
        Optional ByVal compareTypeNames As Boolean _
    ) As Boolean
Attribute RecursiveEvery.VB_Description = "Recursively checks if the passed array only includes the search element."
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
            If Not RecursiveEvery(searchElement, passThru, compareTypeNames:=compareTypeNames) Then
                Exit Function
            End If
        ElseIf compareTypeNames Then
            If Not InStr( _
                    UCase$(TypeName(searchArray(i))), _
                    UCase$(CStr(searchElement)) _
                ) > 0 Then
                Exit Function
            End If
        Else
            If Not ElementsAreEqual(searchElement, searchArray(i)) Then
                Exit Function
            End If
        End If
    Next
    
    RecursiveEvery = True
End Function

'@Description("shuffles the order of the passed array.")
Private Function RecursiveShuffle( _
        ByRef sourceArray() As Variant, _
        Optional ByVal recurse As Boolean _
    ) As Variant()
Attribute RecursiveShuffle.VB_Description = "shuffles the order of the passed array."
    ' https://en.wikipedia.org/wiki/Fisher%E2%80%93Yates_shuffle#The_modern_algorithm
    Dim i As Long
    Dim j As Long
    Dim lower As Long
    Dim upper As Long
    Dim nested() As Variant
    lower = LBound(sourceArray)
    upper = UBound(sourceArray)
    Randomize
    For i = upper To lower + 1 Step -1
        If IsArray(sourceArray(i)) And recurse Then
            nested = sourceArray(i)
            sourceArray(i) = RecursiveShuffle(nested, recurse)
        End If
        j = Int(Rnd * (i - lower) + 1)
        If IsArray(sourceArray(j)) And recurse Then
            nested = sourceArray(j)
            sourceArray(j) = RecursiveShuffle(nested, recurse)
        End If
        Swap sourceArray, i, j
    Next
    RecursiveShuffle = sourceArray
End Function

'@Description("reverses the order of the passed array.")
Private Function RecursiveReverse( _
        ByRef sourceArray() As Variant, _
        Optional ByVal recurse As Boolean _
    ) As Variant()
Attribute RecursiveReverse.VB_Description = "reverses the order of the passed array."
    
    Dim localLength As Long
    Dim middle As Long
    Dim lower As Long
    Dim upper As Long
    Dim passThruArray() As Variant
    
    localLength = GetArrayLength(sourceArray)
    lower = LBound(sourceArray)
    middle = Int(localLength / 2) + lower
    
    Do While lower <> middle
        ' TODO: check this works with any lb
        upper = UBound(sourceArray) + LBound(sourceArray) - lower
        If IsArray(sourceArray(lower)) And recurse Then
            passThruArray = sourceArray(lower)
            sourceArray(lower) = RecursiveReverse(passThruArray, recurse)
        End If
        If IsArray(sourceArray(upper)) And recurse Then
            passThruArray = sourceArray(upper)
            sourceArray(upper) = RecursiveReverse(passThruArray, recurse)
        End If
        Swap sourceArray, lower, upper
        lower = lower + 1
    Loop
    RecursiveReverse = sourceArray
End Function

'@Description("returns the length of the passed array.")
Private Function GetArrayLength(ByRef sourceArray() As Variant) As Long
Attribute GetArrayLength.VB_Description = "returns the length of the passed array."
    GetArrayLength = UBound(sourceArray) - LBound(sourceArray) + 1
End Function

'@Description("converts all the values in a jagged array to a 1d array.")
Private Sub RecursiveFlatten( _
        ByRef sourceArray() As Variant, _
        ByRef results As BetterArray _
    )
Attribute RecursiveFlatten.VB_Description = "converts all the values in a jagged array to a 1d array."
    Dim element As Variant
    For Each element In sourceArray
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
Private Function RecursiveMax(ByRef sourceArray() As Variant) As Variant
Attribute RecursiveMax.VB_Description = "Recursively finds the largest value in an array."
    Dim i As Long
    Dim result As Variant
    
    If IsMultidimensionalArray(sourceArray) Then
        RaiseError EC_INVALID_MULTIDIMENSIONAL_ARRAY_OPERATION, "Max", "Args"
    Else
        For i = LBound(sourceArray) To UBound(sourceArray)
            If IsArray(sourceArray(i)) Then
                Dim nestedArray() As Variant
                Dim nestedResult As Variant
                nestedArray = sourceArray(i)
                nestedResult = RecursiveMax(nestedArray)
                If IsEmpty(result) Then
                    result = nestedResult
                Else
                    If nestedResult > result Then result = nestedResult
                End If
            Else
                Dim currentValue As Variant
                currentValue = GetScalarRepresentation(sourceArray(i))
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
Private Function RecursiveMin(ByRef sourceArray() As Variant) As Variant
Attribute RecursiveMin.VB_Description = "Recursively finds the smallest value in an array."
    Dim i As Long
    Dim result As Variant
    
    If IsMultidimensionalArray(sourceArray) Then
        RaiseError EC_INVALID_MULTIDIMENSIONAL_ARRAY_OPERATION, "Max", "Args"
    Else
        For i = LBound(sourceArray) To UBound(sourceArray)
            If IsArray(sourceArray(i)) Then
                Dim nestedArray() As Variant
                Dim nestedResult As Variant
                nestedArray = sourceArray(i)
                nestedResult = RecursiveMin(nestedArray)
                If IsEmpty(result) Then
                    result = nestedResult
                Else
                    If nestedResult < result Then result = nestedResult
                End If
            Else
                Dim currentValue As Variant
                currentValue = GetScalarRepresentation(sourceArray(i))
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
        Optional ByVal fromIndex As Long, _
        Optional ByVal compareTypeNames As Boolean, _
        Optional ByVal recurse As Boolean _
    ) As Boolean
Attribute RecursiveIncludes.VB_Description = "Recursively checks if the passed array includes the search element."
    Dim localLowerBound As Long
    Dim localUpperBound As Long
    Dim i As Long
    localLowerBound = LBound(searchArray)
    If fromIndex > localLowerBound Then localLowerBound = fromIndex
    localUpperBound = UBound(searchArray)
    For i = localLowerBound To localUpperBound
        If compareTypeNames Then
            If InStr( _
                    UCase$(TypeName(searchArray(i))), _
                    UCase$(CStr(searchElement)) _
                ) > 0 Then
                RecursiveIncludes = True
                Exit Function
            End If
        End If
        If IsArray(searchArray(i)) And recurse Then
            Dim passThru() As Variant
            passThru = searchArray(i)
            If RecursiveIncludes( _
                    searchElement, _
                    passThru, _
                    compareTypeNames:=compareTypeNames, _
                    recurse:=recurse _
                ) Then
                RecursiveIncludes = True
                Exit Function
            End If
        ElseIf Not compareTypeNames Then
            If ElementsAreEqual(searchElement, searchArray(i)) Then
                RecursiveIncludes = True
                Exit Function
            End If
        End If
    Next
    RecursiveIncludes = False
End Function

'@Description("Recursively filters a passed array.")
Private Function RecursiveFilter( _
        ByRef sourceArray() As Variant, _
        ByVal match As Variant, _
        ByVal include As Boolean, _
        ByVal recurse As Boolean, _
        Optional ByVal compareTypeNames As Boolean _
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
    result.LowerBound = this.LowerBound
    For i = localLowerBound To localUpperBound
        If IsArray(sourceArray(i)) And recurse Then
            If IsArrayAllocated(sourceArray(i)) Then
                Dim localItems() As Variant
                localItems = sourceArray(i)
                localItems = RecursiveFilter( _
                    localItems, _
                    match, _
                    include, _
                    recurse, _
                    compareTypeNames _
                )
                ' don't push a returned empty array
                If Not RecursiveEvery(Empty, localItems) Then
                    result.Push localItems
                End If
            Else
                ' push an empty array if current is an unalocated rray
                If compareTypeNames Then
                    isMatch = (InStr( _
                        UCase$(TypeName(sourceArray(i))), _
                        UCase$(CStr(match)) _
                    ) > 0)
                    ' push an empty array if only asked to exclude specific types or type
                    ' matches the requested type (even though it's empty...)
                    If (include And isMatch) Or (Not include And Not isMatch) Then
                        result.Push GetEmptyArray
                    End If
                Else
                    If Not include Then
                        ' push an empty array if only asked to exclude specific values
                        result.Push GetEmptyArray
                    End If
                End If
            End If
        Else
            If compareTypeNames Then
                isMatch = (InStr( _
                    UCase$(TypeName(sourceArray(i))), _
                    UCase$(CStr(match)) _
                ) > 0)
            Else
                isMatch = ElementsAreEqual(match, sourceArray(i))
            End If
            If (include And isMatch) Or (Not include And Not isMatch) Then
                result.Push sourceArray(i)
            End If
        End If
    Next
    RecursiveFilter = result.Items
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
    i = this.LowerBound
    Do
        If currentChar = opener Then
            If i > this.LowerBound Then
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

Private Function UnquoteString(ByRef element As String) As String
    Dim localElement As String
    localElement = element
    If Left$(element, 1) = Chr$(34) And Right$(element, 1) = Chr$(34) Then
        'trim string delimiters
        localElement = Mid$(element, 2, Len(element) - 2)
    End If
    UnquoteString = Trim$(localElement)
End Function

Private Function DuckTypeElement(ByRef element As String) As Variant
    Dim result As Variant
    Dim localElement As String
    localElement = element
    If UCase$(localElement) = "TRUE" Or UCase$(element) = "FALSE" Then
        result = CBool(localElement)
    ElseIf IsNumeric(localElement) Then
        Const Epsilon As Double = 2 ^ -52
        Dim diff As Double
        ' check if integer
        diff = Abs(Fix(localElement) - localElement)
        If diff > Epsilon Then
            result = CDbl(localElement)
        Else
            result = CLng(localElement)
        End If
    Else
        result = UnquoteString(localElement)
    End If
    DuckTypeElement = result
End Function

'@Description("Takes a string array. Returns a Variant array with the elements converted to an appropriate type")
Private Function DuckTypeStringArray(ByRef sourceArray() As String) As Variant()
Attribute DuckTypeStringArray.VB_Description = "Takes a string array. Returns a Variant array with the elements converted to an appropriate type"
    Dim i As Long
    Dim localLowerBound As Long
    Dim localUpperBound As Long
    Dim result() As Variant
    
    localLowerBound = LBound(sourceArray)
    localUpperBound = UBound(sourceArray)
    ReDim result(localLowerBound To localUpperBound)
    For i = localLowerBound To localUpperBound
        result(i) = DuckTypeElement(sourceArray(i))
    Next
    DuckTypeStringArray = result
End Function

'@Description("Populates the ErrorDefinitions array.")
Private Sub PopulateErrorDefinitions()
Attribute PopulateErrorDefinitions.VB_Description = "Populates the ErrorDefinitions array."
    Dim Source As String
    Source = TypeName(Me)
    this.ErrorDefinitions(ErrorCodes.EC_EXPECTED_RANGE_OBJECT) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_EXPECTED_RANGE_OBJECT, _
        Source:=Source, _
        description:="Range Object Expected" _
    )
    this.ErrorDefinitions(ErrorCodes.EC_MAX_DIMENSIONS_LIMIT) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_MAX_DIMENSIONS_LIMIT, _
        Source:=Source, _
        description:="Cannot convert structure of arrays with more than 20 dimensions." _
    )
    this.ErrorDefinitions(ErrorCodes.EC_EXPECTED_COLLECTION_OBJECT) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_EXPECTED_COLLECTION_OBJECT, _
        Source:=Source, _
        description:="Valid Collection Object Expected" _
    )
    this.ErrorDefinitions(ErrorCodes.EC_EXCEEDS_MAX_SORT_DEPTH) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_EXCEEDS_MAX_SORT_DEPTH, _
        Source:=Source, _
        description:="Cannot sort on arrays with more than 2 dimensions" _
    )
    this.ErrorDefinitions(ErrorCodes.EC_EXPECTED_JAGGED_ARRAY) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_EXPECTED_JAGGED_ARRAY, _
        Source:=Source, _
        description:="Expected jagged array." _
    )
    this.ErrorDefinitions(ErrorCodes.EC_EXPECTED_MULTIDIMENSION_ARRAY) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_EXPECTED_MULTIDIMENSION_ARRAY, _
        Source:=Source, _
        description:="Expected multidimension array." _
    )
    this.ErrorDefinitions(ErrorCodes.EC_EXPECTED_ARRAY) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_EXPECTED_ARRAY, _
        Source:=Source, _
        description:="Expected array." _
    )
    this.ErrorDefinitions(ErrorCodes.EC_NULL_STRING) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_NULL_STRING, _
        Source:=Source, _
        description:="Cannot parse from a null string. Expected string with length greater than 0." _
    )
    this.ErrorDefinitions(ErrorCodes.EC_UNALLOCATED_ARRAY) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_UNALLOCATED_ARRAY, _
        Source:=Source, _
        description:="Cannot operate on unallocated array." _
    )
    this.ErrorDefinitions(ErrorCodes.EC_UNDEFINED_ARRAY) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_UNDEFINED_ARRAY, _
        Source:=Source, _
        description:="Array is undefined." _
    )
    this.ErrorDefinitions(ErrorCodes.EC_INVALID_MULTIDIMENSIONAL_ARRAY_OPERATION) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_INVALID_MULTIDIMENSIONAL_ARRAY_OPERATION, _
        Source:=Source, _
        description:="Unable to perform the requested operation on a multidimensional array." _
    )
    this.ErrorDefinitions(ErrorCodes.EC_EXPECTED_VARIANT_ARRAY) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_EXPECTED_VARIANT_ARRAY, _
        Source:=Source, _
        description:="Unable to perform the requested operation on a typed array." _
    )
    this.ErrorDefinitions(ErrorCodes.EC_EXCEEDS_MAX_ARRAY_LENGTH) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_EXCEEDS_MAX_ARRAY_LENGTH, _
        Source:=Source, _
        description:="The requested operation would result in an array which exceeds the maximum possible length." _
    )
    this.ErrorDefinitions(ErrorCodes.EC_STRING_TYPE_EXPECTED) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_STRING_TYPE_EXPECTED, _
        Source:=Source, _
        description:="Expected a String or String-coercible type." _
    )
    this.ErrorDefinitions(ErrorCodes.EC_CANNOT_CONVERT_TO_REQUESTED_STRUCTURE) = ErrorDefinitionFactory( _
        number:=ErrorCodes.EC_CANNOT_CONVERT_TO_REQUESTED_STRUCTURE, _
        Source:=Source, _
        description:="The stored array cannot be converted to the requested structure." _
    )
    
End Sub

'@Description("Returns an ErrorDefinition Type populated with the provided arguments.")
Private Function ErrorDefinitionFactory( _
        ByVal number As Long, _
        ByVal Source As String, _
        ByVal description As String _
    ) As ErrorDefinition
Attribute ErrorDefinitionFactory.VB_Description = "Returns an ErrorDefinition Type populated with the provided arguments."
    Dim result As ErrorDefinition
    result.number = number
    result.Source = Source
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
    currentError = this.ErrorDefinitions(errorCode)
    Err.Raise currentError.number, _
              currentError.Source & "." & caller, _
              localArgName & currentError.description
End Sub

'Used by ToExcelRange
'@Description("Ensures each element of passed array can be represented as a scalar value.")
Private Function EnsureScalar1DArray(ByRef sourceArray() As Variant) As Variant()
Attribute EnsureScalar1DArray.VB_Description = "Ensures each element of passed array can be represented as a scalar value."
    Dim i As Long
    Dim localLowerBound As Long
    Dim localUpperBound As Long
    Dim result() As Variant
    
    localLowerBound = LBound(sourceArray)
    localUpperBound = UBound(sourceArray)
    ReDim result(localLowerBound To localUpperBound)
    For i = localLowerBound To localUpperBound
        result(i) = GetScalarRepresentation(sourceArray(i))
    Next
    EnsureScalar1DArray = result
End Function

'@Description("Converts a 1d array to a 1 column wide 2d array.")
Private Function Transpose1DArray(ByRef sourceArray() As Variant) As Variant()
Attribute Transpose1DArray.VB_Description = "Converts a 1d array to a 1 column wide 2d array."
    Dim i As Long
    Dim j As Long
    Dim localUpperBound As Long
    Dim localLowerBound As Long
    Dim result() As Variant
    localUpperBound = UBound(sourceArray)
    localLowerBound = LBound(sourceArray)
    ReDim result(localLowerBound To localUpperBound, _
                 localLowerBound To localLowerBound)
    j = localLowerBound
    For i = localLowerBound To localUpperBound
        result(j, localLowerBound) = sourceArray(i)
        j = j + 1
    Next
    Transpose1DArray = result
End Function

'@Description("Transposes the values in a 2d array. Rows become columns, columns become rows.")
Private Function Transpose2DArray(ByRef sourceArray() As Variant) As Variant()
Attribute Transpose2DArray.VB_Description = "Transposes the values in a 2d array. Rows become columns, columns become rows."
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
        For currentColumn = LowerBoundCol To UpperBoundCol
            result(currentRow, currentColumn) = sourceArray(currentColumn, currentRow)
        Next
    Next
    Transpose2DArray = result
End Function

'@Description("Transposes the values in a jagged array with a depth of 2. Rows become columns, columns become rows.")
Private Function TransposeArrayOfArrays(ByRef sourceArray() As Variant) As Variant()
Attribute TransposeArrayOfArrays.VB_Description = "Transposes the values in a jagged array with a depth of 2. Rows become columns, columns become rows."
    Dim currentRow As Long
    Dim LowerBoundRow As Long
    Dim UpperBoundRow As Long
    Dim currentColumn As Long
    Dim LowerBoundCol As Long
    Dim UpperBoundCol As Long
    Dim result() As Variant
    Dim nestedBounds() As Long
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
    TransposeArrayOfArrays = result
End Function

'@Description("Returns an empty 1d array with 1 slot.")
Private Function GetEmptyArray() As Variant()
Attribute GetEmptyArray.VB_Description = "Returns an empty 1d array with 1 slot."
    Dim result() As Variant
    ReDim result(this.LowerBound To this.LowerBound)
    GetEmptyArray = result
End Function

'@Description("Sorts in place a passed array using the QuickSort algorithm.")
Private Sub QuickSort( _
        ByRef sourceArray() As Variant, _
        ByVal low As Long, _
        ByVal high As Long, _
        ByVal localArrayType As ArrayTypes, _
        Optional ByVal col As Long _
    )
Attribute QuickSort.VB_Description = "Sorts in place a passed array using the QuickSort algorithm."
    Dim partitionIndex As Long
    If low < high Then
        partitionIndex = QsPartition(sourceArray, low, high, localArrayType, col)
        QuickSort sourceArray, low, partitionIndex - 1, localArrayType, col
        QuickSort sourceArray, partitionIndex + 1, high, localArrayType, col
    End If
End Sub

'@Description("Child function of quickSort. Returns the next partition index.")
Private Function QsPartition( _
        ByRef sourceArray() As Variant, _
        ByVal low As Long, _
        ByVal high As Long, _
        ByVal localArrayType As ArrayTypes, _
        Optional ByVal col As Long _
    ) As Long
Attribute QsPartition.VB_Description = "Child function of quickSort. Returns the next partition index."
    Dim i As Long
    Dim j As Long
    Dim pivot As Variant
    Dim current As Variant
    
    If localArrayType = BA_JAGGED Then
        pivot = sourceArray(high)(col)
    Else
        pivot = sourceArray(high)
    End If
    
    i = low - 1
    For j = low To high - 1
        If localArrayType = BA_JAGGED Then
            current = sourceArray(j)(col)
        Else
            current = sourceArray(j)
        End If
    
        If current <= pivot Then
            i = i + 1
            Swap sourceArray, i, j
        End If
    Next
    Swap sourceArray, i + 1, high
    QsPartition = i + 1
End Function

'@Description("Swaps the location of two elements in an array.")
Private Sub Swap(ByRef sourceArray() As Variant, ByVal i As Long, ByVal j As Long)
Attribute Swap.VB_Description = "Swaps the location of two elements in an array."
    Dim element As Variant
    LetOrSetElement element, sourceArray(i)
    LetOrSetElement sourceArray(i), sourceArray(j)
    '@Ignore UnassignedVariableUsage
    LetOrSetElement sourceArray(j), element
End Sub

'@Description("Sets or lets the value of a variable by reference depending on the type of the element to be assigned")
Private Sub LetOrSetElement(ByRef destination As Variant, ByRef Source As Variant)
Attribute LetOrSetElement.VB_Description = "Sets or lets the value of a variable by reference depending on the type of the element to be assigned"
    If IsObject(Source) Then
        Set destination = Source
    Else
        destination = Source
    End If
End Sub

'@Description("Efficiently concatenates string fragments to a byte array buffer")
Private Function StringBuilder(Optional ByVal Fragment As String = vbNullString, Optional ByVal Final As Boolean, Optional ByVal NewString As Boolean) As String
Attribute StringBuilder.VB_Description = "Efficiently concatenates string fragments to a byte array buffer"
    Static buffer() As Byte
    Static bufferLength As Long
    Static bufferCapacity As Long
    Static cursor As Long
    Static firstIndex As Long
    
    Dim fragLength As Long
    Dim lastIndex As Long
    Dim i As Long

    fragLength = LenB(Fragment)
    If cursor < 1 Or NewString Then
        buffer = Fragment
        bufferLength = fragLength
        bufferCapacity = fragLength
        firstIndex = LBound(buffer)
        cursor = UBound(buffer)
    ElseIf fragLength > 0 Then
        bufferLength = bufferLength + fragLength
        lastIndex = firstIndex + bufferLength - 1
        If bufferLength > bufferCapacity Then
            Do
                bufferCapacity = bufferCapacity * 2
            Loop While bufferLength > bufferCapacity
            ReDim Preserve buffer(firstIndex To firstIndex + bufferCapacity)
        End If
        For i = cursor + 1 To lastIndex
            buffer(i) = AscB(MidB$(Fragment, i - cursor, 1))
        Next
        cursor = lastIndex
    End If
    
    If Final Then
        ReDim Preserve buffer(firstIndex To cursor)
        StringBuilder = CStr(buffer)
        Erase buffer
        cursor = Empty
    End If
    
End Function

'@Description("Recursively parses an array to a string representation.")
Private Sub RecursiveToString( _
        ByRef sourceArray() As Variant, _
        ByVal prettyPrint As Boolean, _
        ByVal separator As String, _
        ByVal openingDelimiter As String, _
        ByVal closingDelimiter As String, _
        ByVal quoteStrings As Boolean, _
        Optional ByVal tabs As Long = 1 _
    )
Attribute RecursiveToString.VB_Description = "Recursively parses an array to a string representation."
    Const tabWidth As Long = 2
    Dim i As Long
    Dim sep As String
    
    StringBuilder openingDelimiter
    For i = LBound(sourceArray) To UBound(sourceArray)
        sep = IIf(i = UBound(sourceArray), closingDelimiter, separator)
        If IsArray(sourceArray(i)) Then
            Dim nested() As Variant
            nested = sourceArray(i)
            If prettyPrint Then
                StringBuilder vbCrLf & Space(tabWidth * tabs)
            End If
            RecursiveToString nested, prettyPrint, separator, openingDelimiter, closingDelimiter, quoteStrings, tabs + 1
            If i = UBound(sourceArray) And prettyPrint Then
                StringBuilder vbCrLf
                StringBuilder Space(tabWidth * (tabs - 1))
            End If
            StringBuilder sep
        Else
            ' ensure no arrays passed to GetScalar as it clears string builder
            StringBuilder Replace( _
                    CStr(GetScalarRepresentation(sourceArray(i), quoteStrings)), _
                    ",", _
                    vbNullString _
                ) & sep
        End If
    Next
End Sub

'@Description("Returns scalar values as is or appropriate representation if not scalar.")
Private Function GetScalarRepresentation( _
        ByRef element As Variant, _
        Optional ByVal quoteStrings As Boolean _
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
        StringBuilder NewString:=True
        RecursiveToString passThruArray, False, ",", "{", "}", quoteStrings
        result = StringBuilder(Final:=True)
    End If
    If Not IsNumeric(result) And quoteStrings Then
        GetScalarRepresentation = Chr$(34) & result & Chr$(34)
    Else
        GetScalarRepresentation = result
    End If
    
End Function

'@Description("Returns the type of a passed array.")
Private Function GetArrayType(ByVal sourceArray As Variant) As ArrayTypes
Attribute GetArrayType.VB_Description = "Returns the type of a passed array."
    Dim result As ArrayTypes
    If IsArray(sourceArray) Then
        If Not IsArrayAllocated(sourceArray) Then
            result = BA_UNALLOCATED
        Else
            If IsMultidimensionalArray(sourceArray) Then
                result = BA_MULTIDIMENSION
            ElseIf IsJaggedArray(sourceArray) Then
                result = BA_JAGGED
            Else
                result = BA_ONEDIMENSION
            End If
        End If
    Else
        If IsEmpty(sourceArray) Then
            result = BA_UNDEFINED
        Else
            RaiseError EC_EXPECTED_ARRAY, "getArrayType", "sourceArray"
        End If
    End If
    GetArrayType = result
End Function

'@Description("Tests if an array has been allocated.")
Private Function IsArrayAllocated(ByVal sourceArray As Variant) As Boolean
Attribute IsArrayAllocated.VB_Description = "Tests if an array has been allocated."
    On Error Resume Next
    IsArrayAllocated = ( _
        IsArray(sourceArray) And _
        Not IsError(LBound(sourceArray, 1)) And _
        LBound(sourceArray, 1) <= UBound(sourceArray, 1) _
    )
    On Error GoTo 0
End Function

'@Description("Tests if an array is jagged.")
Private Function IsJaggedArray(ByVal sourceArray As Variant) As Boolean
Attribute IsJaggedArray.VB_Description = "Tests if an array is jagged."
    If IsArray(sourceArray) Then
        On Error GoTo ErrHandler
        Dim element As Variant
        For Each element In sourceArray
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
Private Function IsMultidimensionalArray(ByVal sourceArray As Variant) As Boolean
Attribute IsMultidimensionalArray.VB_Description = "Tests if an array is multidimensioned."
    If IsArray(sourceArray) Then
        On Error GoTo ErrHandler
        '@Ignore VariableNotUsed
        Dim localUpperBound As Long
        '@Ignore AssignmentNotUsed
        localUpperBound = UBound(sourceArray, 2)
        IsMultidimensionalArray = True
        On Error GoTo 0
    End If
    Exit Function
ErrHandler:
    Err.Clear
End Function

'@Description("re-indexes array to conform to new bounds. Does not change length.")
Private Function Rebase(Optional ByRef sourceArray As Variant, Optional ByVal CurrentArrayType As ArrayTypes) As Variant()
Attribute Rebase.VB_Description = "re-indexes array to conform to new bounds. Does not change length."
    Dim localItems() As Variant
    Dim localType As ArrayTypes

    If IsMissing(sourceArray) Or Not IsArray(sourceArray) Then
        localItems = InternalItems
        localType = this.ArrayType
    Else
        localItems = sourceArray
        If CurrentArrayType = 0 Then
            localType = GetArrayType(localItems)
        Else
            localType = CurrentArrayType
        End If
    End If
    If localType = BA_MULTIDIMENSION Then
        ' ensure all nested arrays have the correct base
        ' if they don't and additional arrays are added via concat the resulting
        ' md array from jaggedToMulti will have unexpected bounds
        ' NOTE: This may need to be applied to jagged too for consistency
        Rebase = RecursiveRebase(localItems, True)
    Else
        Rebase = RecursiveRebase(localItems, False)
    End If
End Function

'@Description("re-indexes nested arrays to conform to new bounds. Child function of Rebase)
Private Function RecursiveRebase( _
        ByRef sourceArray() As Variant, _
        ByVal recurse As Boolean _
    ) As Variant()
    Dim i As Long
    Dim localLowerBound As Long
    Dim localUpperBound As Long
    Dim offset As Long
    Dim newItems() As Variant
    
    localLowerBound = LBound(sourceArray)
    localUpperBound = UBound(sourceArray)
    offset = this.LowerBound - localLowerBound
    ReDim newItems(this.LowerBound To localUpperBound + offset)
    For i = localLowerBound To localUpperBound
        If IsArray(sourceArray(i)) And recurse And offset <> 0 Then
            Dim nested() As Variant
            nested = sourceArray(i)
            newItems(i + offset) = RecursiveRebase(nested, recurse)
        Else
            newItems(i + offset) = sourceArray(i)
        End If
    Next
    
    RecursiveRebase = newItems
End Function

'@Description("Increases capacity of internal array if required.")
Private Sub EnsureCapacity(ByVal MinimumCapacity As Long)
Attribute EnsureCapacity.VB_Description = "Increases capacity of internal array if required."
    If this.Capacity < MinimumCapacity Then
        Dim newCapacity As Long
        newCapacity = IIf(this.Capacity = 0, DEFAULT_CAPACITY, this.Capacity * 2)
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
        If Not IsJaggedArray(sourceArray) Then
            RaiseError EC_EXPECTED_JAGGED_ARRAY, "JaggedToMulti", "sourceArray()"
            Exit Function
        End If
        localDepth = GetJaggedArrayDepth(sourceArray)
    End If
    
    If localDepth <= 1 Then
        result = sourceArray
    ElseIf localDepth = 2 Then
        Dim firstDimBounds() As Long
        Dim secondDimBounds() As Long
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
        result = RecursiveJaggedToMulti(sourceArray, localDepth, ensureScalar:=ensureScalar)
    Else
        RaiseError EC_MAX_DIMENSIONS_LIMIT, "JaggedToMulti()", "sourceArray()"
        result = sourceArray
    End If
    JaggedToMulti = result
End Function

Private Function MultiToJagged( _
        ByRef sourceArray() As Variant, _
        Optional ByVal depth As Long _
    ) As Variant()
    Dim localDepth As Long
    Dim result() As Variant
    
    localDepth = depth
    
    If localDepth = 0 Then
        If Not IsMultidimensionalArray(sourceArray) Then
            RaiseError EC_EXPECTED_MULTIDIMENSION_ARRAY, "MultiToJagged", "sourceArray()"
            result = sourceArray
            Exit Function
        End If
        localDepth = GetMultidimensionalArrayDepth(sourceArray)
    End If
    
    If localDepth <= 1 Then
        result = sourceArray
    ElseIf localDepth = 2 Then
        Dim firstDimBounds(0 To 1) As Long
        Dim secondDimBounds(0 To 1) As Long
        Dim firstDimOffset As Long
        Dim secondDimOffset As Long
        Dim nested() As Variant
        Dim i As Long
        Dim j As Long
        
        firstDimBounds(0) = LBound(sourceArray, 1)
        firstDimBounds(1) = UBound(sourceArray, 1)
        secondDimBounds(0) = LBound(sourceArray, 2)
        secondDimBounds(1) = UBound(sourceArray, 2)
        firstDimOffset = this.LowerBound - firstDimBounds(0)
        secondDimOffset = this.LowerBound - secondDimBounds(0)
        ReDim result(firstDimBounds(0) + firstDimOffset To firstDimBounds(1) + firstDimOffset)
        For i = firstDimBounds(0) To firstDimBounds(1)
            ReDim nested(secondDimBounds(0) + secondDimOffset To secondDimBounds(1) + secondDimOffset)
            For j = secondDimBounds(0) To secondDimBounds(1)
                LetOrSetElement nested(j + secondDimOffset), sourceArray(i, j)
            Next
            result(i + firstDimOffset) = nested
        Next
    ElseIf localDepth > 2 And localDepth <= 20 Then
        result = RecursiveMultiToJagged(sourceArray, localDepth)
    Else
        RaiseError EC_MAX_DIMENSIONS_LIMIT, "MultiToJagged()", "sourceArray()"
        result = sourceArray
    End If
    
    MultiToJagged = result
End Function

'@Description("Recursively Converts multidimension arrays to jagged arrays. For md arrays with more than 2d.")
Private Function RecursiveMultiToJagged( _
        ByRef sourceArray() As Variant, _
        Optional ByVal depth As Long, _
        Optional ByVal currentDepth As Long, _
        Optional ByRef crumbs As Variant _
    ) As Variant()
Attribute RecursiveMultiToJagged.VB_Description = "Recursively Converts multidimension arrays to jagged arrays. For md arrays with more than 2d."
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
    
    If localDepth > 1 Then
        If IsMissing(crumbs) Then
            ReDim localCrumbs(localDepth - 1)
        Else
            localCrumbs = crumbs
        End If
        localCurrentDepth = localCurrentDepth + 1
        localLowerBound = LBound(sourceArray, localCurrentDepth)
        localUpperBound = UBound(sourceArray, localCurrentDepth)
        boundOffset = this.LowerBound - localLowerBound
        ReDim result(localLowerBound + boundOffset To localUpperBound + boundOffset)
        For i = localLowerBound To localUpperBound
            localCrumbs(localCurrentDepth - 1) = i
            If localCurrentDepth = localDepth Then
                LetOrSetElement result(i + boundOffset), GetElementByBreadcrumb(sourceArray, localCrumbs)
            Else
                result(i + boundOffset) = RecursiveMultiToJagged(sourceArray, localDepth, localCurrentDepth, localCrumbs)
            End If
        Next
    Else
        result = sourceArray
    End If
    RecursiveMultiToJagged = result
End Function


'@Description("Recursively Converts jagged arrays to multidimension arrays. For jagged arrays with more than 2d.")
Private Function RecursiveJaggedToMulti( _
        ByRef sourceArray() As Variant, _
        Optional ByVal depth As Long, _
        Optional ByVal currentDepth As Long, _
        Optional ByRef crumbs As Variant, _
        Optional ByRef result As Variant, _
        Optional ByVal ensureScalar As Boolean _
    ) As Variant()
Attribute RecursiveJaggedToMulti.VB_Description = "Recursively Converts jagged arrays to multidimension arrays. For jagged arrays with more than 2d."
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
        localResult = CreateMultidimensionalArray(MapJaggedArray(sourceArray, KnownDepth:=localDepth))
    Else
        localResult = result
    End If

    localCurrentDepth = localCurrentDepth + 1
    localLowerBound = LBound(sourceArray)
    localUpperBound = UBound(sourceArray)
    For i = localLowerBound To localUpperBound
        localCrumbs(localCurrentDepth - 1) = i
        If localCurrentDepth = localDepth Then
            If ensureScalar Then
                localResult = LetElementByBreadcrumb(localResult, localCrumbs, GetScalarRepresentation(sourceArray(i)))
            Else
                localResult = LetElementByBreadcrumb(localResult, localCrumbs, sourceArray(i))
            End If
        Else
            Dim nested() As Variant
            If Not IsArray(sourceArray(i)) Then
                nested = Array(sourceArray(i))
            Else
                nested = sourceArray(i)
            End If
            localResult = RecursiveJaggedToMulti( _
                sourceArray:=nested, _
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

'@Description("Returns a number representing how many levels of nested arrays are contained in array.")
Private Function GetJaggedArrayDepth(ByRef sourceArray() As Variant) As Long
Attribute GetJaggedArrayDepth.VB_Description = "Returns a number representing how many levels of nested arrays are contained in array."
    Dim i As Long
    Dim localLowerBound As Long
    Dim localUpperBound As Long
    Dim currentDepth As Long
    Dim maxDepth As Long
    Dim depth As Long
    
    If IsArray(sourceArray) Then
        depth = depth + 1
        localLowerBound = LBound(sourceArray)
        localUpperBound = UBound(sourceArray)
        For i = localLowerBound To localUpperBound
            If IsArray(sourceArray(i)) Then
                Dim nested() As Variant
                nested = sourceArray(i)
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
    ) As Long()
Attribute GetMaxBoundsAtDimension.VB_Description = "Gets the maximium bound values for a specified dimension of a jagged array."
    Dim maxBounds() As Long
    
    If currentDimension < dimension Then
        Dim i As Long
        Dim currentBounds() As Long
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
Private Function GetArrayBounds(ByVal sourceArray As Variant) As Long()
Attribute GetArrayBounds.VB_Description = "Returns a 0-based array with two slots. First slot is the lowerbound of the passed array. second slot (index 0) is the upper bound."
    Dim result(0 To 1) As Long
    If IsArray(sourceArray) Then
        result(0) = LBound(sourceArray)
        result(1) = UBound(sourceArray)
    End If
    GetArrayBounds = result
End Function

'@Description("Returns a number representing the total number of dimensions of a multidimension array.")
Private Function GetMultidimensionalArrayDepth(ByVal sourceArray As Variant) As Long
Attribute GetMultidimensionalArrayDepth.VB_Description = "Returns a number representing the total number of dimensions of a multidimension array."
    Dim i As Long
    '@Ignore VariableNotUsed
    Dim void As Long
    
    On Error Resume Next
    Do
        i = i + 1
        void = UBound(sourceArray, i)
    Loop Until Err.number <> 0
    Err.Clear
    On Error GoTo 0
    GetMultidimensionalArrayDepth = i - 1
End Function

'@Description("Returns the element stored in an array based on a crumb argument. A crumb is an array representing the location of the multidimension array slot.")
Private Function GetElementByBreadcrumb( _
        ByRef sourceArray() As Variant, _
        ByRef crumb() As Variant _
    ) As Variant
Attribute GetElementByBreadcrumb.VB_Description = "Returns the element stored in an array based on a crumb argument. A crumb is an array representing the location of the multidimension array slot."
    Dim result As Variant
    Select Case UBound(crumb)
    Case 0
        LetOrSetElement result, sourceArray(crumb(0))
    Case 1
        LetOrSetElement result, sourceArray(crumb(0), crumb(1))
    Case 2
        LetOrSetElement result, sourceArray(crumb(0), crumb(1), crumb(2))
    Case 3
        LetOrSetElement result, sourceArray(crumb(0), crumb(1), crumb(2), crumb(3))
    Case 4
        LetOrSetElement result, sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4))
    Case 5
        LetOrSetElement result, sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5))
    Case 6
        LetOrSetElement result, sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6))
    Case 7
        LetOrSetElement result, sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7))
    Case 8
        LetOrSetElement result, sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8))
    Case 9
        LetOrSetElement result, sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9))
    Case 10
        LetOrSetElement result, sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10))
    Case 11
        LetOrSetElement result, sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10), crumb(11))
    Case 12
        LetOrSetElement result, sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10), crumb(11), crumb(12))
    Case 13
        LetOrSetElement result, sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10), crumb(11), crumb(12), crumb(13))
    Case 14
        LetOrSetElement result, sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10), crumb(11), crumb(12), crumb(13), crumb(14))
    Case 15
        LetOrSetElement result, sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15))
    Case 16
        LetOrSetElement result, sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16))
    Case 17
        LetOrSetElement result, sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17))
    Case 18
        LetOrSetElement result, sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18))
    Case 19
        LetOrSetElement result, sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18), crumb(19))
    Case 20
        LetOrSetElement result, sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), _
            crumb(9), crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18), crumb(19), crumb(20))
    End Select
    '@Ignore UnassignedVariableUsage
    GetElementByBreadcrumb = result
End Function

'@Description("Assigns an element to a location in a multidimensonal array based on a crumb argument. A crumb is an array representing the location of the multidimension array slot.")
Private Function LetElementByBreadcrumb( _
        ByRef sourceArray() As Variant, _
        ByRef crumb() As Variant, _
        ByVal element As Variant _
    ) As Variant
Attribute LetElementByBreadcrumb.VB_Description = "Assigns an element to a location in a multidimensonal array based on a crumb argument. A crumb is an array representing the location of the multidimension array slot."
    Select Case UBound(crumb)
    Case 0
        LetOrSetElement sourceArray(crumb(0)), element
    Case 1
        LetOrSetElement sourceArray(crumb(0), crumb(1)), element
    Case 2
        LetOrSetElement sourceArray(crumb(0), crumb(1), crumb(2)), element
    Case 3
        LetOrSetElement sourceArray(crumb(0), crumb(1), crumb(2), crumb(3)), element
    Case 4
        LetOrSetElement sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4)), element
    Case 5
        LetOrSetElement sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5)), element
    Case 6
        LetOrSetElement sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6)), element
    Case 7
        LetOrSetElement sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7)), element
    Case 8
        LetOrSetElement sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8)), element
    Case 9
        LetOrSetElement sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9)), _
            element
    Case 10
        LetOrSetElement sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
            crumb(10)), element
    Case 11
        LetOrSetElement sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
            crumb(10), crumb(11)), element
    Case 12
        LetOrSetElement sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
            crumb(10), crumb(11), crumb(12)), element
    Case 13
        LetOrSetElement sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
            crumb(10), crumb(11), crumb(12), crumb(13)), element
    Case 14
        LetOrSetElement sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
            crumb(10), crumb(11), crumb(12), crumb(13), crumb(14)), element
    Case 15
        LetOrSetElement sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
            crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15)), element
    Case 16
        LetOrSetElement sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
        crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16)), element
    Case 17
        LetOrSetElement sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
            crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17)), element
    Case 18
        LetOrSetElement sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
            crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18)), element
    Case 19
        LetOrSetElement sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
        crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18), crumb(19)), element
    Case 20
        LetOrSetElement sourceArray(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), _
        crumb(10), crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18), crumb(19), crumb(20)), _
        element
    End Select

    LetElementByBreadcrumb = sourceArray
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
