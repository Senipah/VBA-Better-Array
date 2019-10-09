VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DynamicArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@Folder("DynamicArray")

'@IgnoreModule MultipleDeclarations
'@IgnoreModule ProcedureNotUsed
'@IgnoreModule FunctionReturnValueNotUsed

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

''''''''''
' Fields '
''''''''''

Private pCapacity As Long
Private pLength As Long
Private pBase As Long
Private pEmptyArray(0) As Variant
Private pItems() As Variant
Private pArrayType As ArrayTypes

'''''''''''''''
' Constructor '
'''''''''''''''

Private Sub Class_Initialize()
    Me.Capacity = DEFAULT_CAPACITY
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

Public Property Let Capacity(ByVal value As Long)
    If value < 0 Then Err.Raise 9
    If value <> pCapacity Then
        If pCapacity > 0 Then
            If (UBound(pItems) - LBound(pItems) + 1) <> value Then
                Dim newItems() As Variant
                newItems = pItems
                ReDim Preserve newItems(pBase To (value + pBase - 1))
                pItems = newItems
            End If
        Else
            ReDim pItems(pBase To (DEFAULT_CAPACITY + pBase - 1))
        End If
        pCapacity = UBound(pItems) - pBase + 1
    End If
End Property

Public Property Get Capacity() As Long
    If IsEmpty(pCapacity) Then Me.Capacity = DEFAULT_CAPACITY
    Capacity = pCapacity
End Property

Public Property Get Length() As Long
    Length = pLength
End Property

Public Property Get Items()
    Dim result() As Variant
    result = getInternalItems()
    If pArrayType = ArrayTypes.MultiDimension And _
       getArrayType(result) = ArrayTypes.Jagged Then result = jaggedToMulti(result)
    Items = result
End Property

Public Property Let Items(ByVal Values As Variant)
    Dim localBase As Long
    Dim localValues As Variant
    localValues = Values
    
    pArrayType = getArrayType(localValues)
    If pArrayType = ArrayTypes.Undefined _
       Or pArrayType = ArrayTypes.Unallocated Then
        Err.Raise 9
    Else
        If pArrayType = ArrayTypes.MultiDimension Then
            localValues = multiToJagged(localValues)
        End If
        localBase = LBound(localValues)
        pItems = localValues
        pLength = UBound(pItems) - LBound(pItems) + 1
        If localBase <> pBase Then
            Rebase
        End If
        Me.Capacity = pLength
    End If
End Property

Public Property Get UpperBound()
    UpperBound = IIf(pLength = 0, pLength, pLength + pBase - 1)
End Property

Public Property Get Base() As Long
    Base = pBase
End Property

Public Property Let Base(ByVal value As Long)
    If value <> pBase And value >= 0 Then
        pBase = value
        Rebase
    End If
End Property

Public Property Let Item(ByVal Index As Long, ByVal Element As Variant)
    If Me.UpperBound >= Index Then
        If Index < Me.Base Then
            Me.Unshift Element
        Else
            If IsObject(Element) Then
                Set pItems(Index) = Element
            Else
                pItems(Index) = Element
            End If
        End If
    Else
        Me.Push Element
    End If
End Property

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

''''''''''''''''''''''
' Private Properties '
''''''''''''''''''''''

''''''''''''''''''
' Public Methods '
''''''''''''''''''

Public Function Push(ParamArray Args() As Variant) As Variant
    Dim Element As Variant
    If pArrayType = ArrayTypes.Unallocated Or _
       pArrayType = ArrayTypes.Undefined Then
        pArrayType = ArrayTypes.OneDimension
    End If
    For Each Element In Args
        If IsArray(Element) And pArrayType = ArrayTypes.OneDimension Then
            pArrayType = ArrayTypes.Jagged
        End If
        If pLength = pCapacity Then
            EnsureCapacity pLength + 1
        End If
        If IsObject(Element) Then
            Set pItems(pLength) = Element
        Else
            pItems(pLength) = Element
        End If
        pLength = pLength + 1
    Next
    Push = pLength
End Function

Public Function Pop() As Variant
    Dim newItems() As Variant
    Dim result As Variant

    If pLength > 0 Then
        result = pItems(Me.UpperBound)
        pLength = pLength - 1
        ReDim newItems(pBase To Me.UpperBound)
        Me.Items = copyRange(newItems)
    End If
    Pop = result
End Function

Public Function Shift() As Variant
    Dim result As Variant
    Dim newItems() As Variant
    If pLength > 0 Then
        result = pItems(pBase)
        ReDim newItems((pBase + 1) To Me.UpperBound)
        Me.Items = copyRange(newItems)
    End If
    Shift = result
End Function

Public Function Unshift(ParamArray Args() As Variant) As Long
    Dim newItems() As Variant
    Dim oldItems() As Variant
    Dim oldType As ArrayTypes
    newItems = Args
    oldType = pArrayType
    oldItems = getInternalItems()
    Me.Items = newItems
    If oldType <> Unallocated And oldType <> Undefined Then
        Me.Concat oldItems
        pArrayType = oldType
    End If
    Unshift = Me.Length
End Function

Public Function Concat(ParamArray Args() As Variant) As Variant
    Dim arg As Variant, Element As Variant
    For Each arg In Args
        If IsArray(arg) Then
            For Each Element In arg
                Me.Push Element
            Next
        Else
            Me.Push arg
        End If
    Next
    Concat = Me.Items
End Function

Public Function CopyFromCollection(ByVal SourceCollection As Collection) As Variant
    If SourceCollection Is Nothing Then Err.Raise 448
    Dim i As Long
    Dim newItems() As Variant
    pLength = SourceCollection.count
    If pLength = 0 Then
        newItems = pEmptyArray
    Else
        ReDim newItems(pBase To (pLength - pBase - 1))
        For i = 1 To pLength
            newItems(i + pBase - 1) = SourceCollection.Item(i)
        Next
    End If
    Me.Items = newItems
    CopyFromCollection = newItems
End Function

Public Function ToString(Optional ByVal PrettyPrint As Boolean) As String
    Dim ArrayType As ArrayTypes
    Dim result As String
    Dim localItems As Variant
    
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

Public Function Sort(Optional ByVal SortColumn As Long) As Variant
    Dim localItems As Variant, sortedItems As Variant
    Dim ArrayType As ArrayTypes
    
    localItems = getInternalItems()
    ArrayType = getArrayType(localItems)
    If ArrayType = ArrayTypes.Unallocated Or _
       ArrayType = ArrayTypes.Undefined Then
        sortedItems = pEmptyArray
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
    Sort = sortedItems
End Function

Public Function CopyWithin(ByVal Target As Long, Optional ByVal StartIndex As Long, Optional ByVal EndIndex As Long) As Variant
    Dim localLength As Long, relativeTarget As Long, relativeStart As Long, relativeEnd As Long
    Dim idx As Long, from As Long, final As Long, count As Long, direction As Long
    Dim arr As Variant
    
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
    CopyWithin = arr
End Function

'@Ignore ShadowedDeclaration
' TODO: add support for 2d arrays.
Public Function Filter(ByVal Match As Variant, Optional ByVal Exclude As Boolean) As Variant
    Dim Element As Variant
    Dim oldArr() As Variant, newArr() As Variant
    oldArr = getInternalItems()
    ReDim newArr(LBound(oldArr), UBound(oldArr))
    Dim i As Long
    i = Me.Base
    For Each Element In oldArr
        If (Exclude And Element = Match) Or (Not Exclude And Element <> Match) Then
            newArr(i) = Element
            i = i + 1
        End If
    Next
    ReDim Preserve newArr(Me.Base, i)
    Me.Items = newArr
    Filter = newArr
End Function

Public Function Includes(ByVal SearchElement As Variant, Optional ByVal FromIndex As Long) As Boolean
    Dim localLength As Long
    Dim currentIndex As Long
    
    Dim arr As Variant
    
    arr = getInternalItems()
    localLength = Me.Length
    
    If localLength = 0 Then
        Includes = False
        Exit Function
    End If
        
    If FromIndex > Me.Base Then
        currentIndex = FromIndex
    ElseIf FromIndex = 0 Then
        'assume 0 val should always be from beginning - cover cases where val is missing
        currentIndex = Me.Base
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

Public Function Keys() As Variant
    Dim i As Long, localBase As Long, localUpperBound As Long
    localBase = Me.Base
    localUpperBound = Me.UpperBound
    Dim result() As Variant: ReDim result(localBase To localUpperBound)
    For i = localBase To localUpperBound: result(i) = i: Next
    Keys = result
End Function

Private Function Max(ParamArray Args() As Variant) As Variant
    Dim i As Long
    Dim arr() As Variant
    Dim result As Variant
    
    If UBound(Args) < LBound(Args) Then          'Test if no args provided
        arr = getInternalItems()
    Else
        arr = Args
        If IsArray(arr) Then
            ' exit function if array is multi dimensional
            If getArrayType(arr) = MultiDimension Then
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

Private Function Min(ParamArray Args() As Variant) As Variant
    Dim i As Long
    Dim arr() As Variant
    Dim result As Variant
    
    If UBound(Args) < LBound(Args) Then          'Test if no args provided
        arr = getInternalItems()
    Else
        arr = Args
        If IsArray(arr) Then
            ' exit function if array is multi dimensional
            If getArrayType(arr) = MultiDimension Then
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

Public Function Slice(ByVal StartIndex As Long, Optional ByVal EndIndex As Long) As Variant
    '@Ignore ShadowedDeclaration
    Dim localLength As Long, relativeStart As Long, relativeEnd As Long
    Dim oldIndex As Long, final As Long, count As Long, newIndex As Long
    Dim arr() As Variant, result() As Variant
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

Public Function Reverse() As Variant
    Dim i As Long, localUpperBound As Long
    Dim arr() As Variant, result() As Variant
    arr = getInternalItems()
    localUpperBound = Me.UpperBound
    ReDim result(Me.Base To localUpperBound)
    For i = localUpperBound To Me.Base Step -1
        result = arr(i)
    Next
    Me.Items = result
    Reverse = result
End Function

'''''''''''''''''''
' Private Methods '
'''''''''''''''''''

Private Sub quickSort(ByRef arr As Variant, ByVal low As Long, ByVal high As Long, ByVal ArrayType As ArrayTypes, Optional ByVal col As Long)
    Dim partitionIndex As Long
    If low < high Then
        partitionIndex = qsPartition(arr, low, high, ArrayType, col)
        quickSort arr, low, partitionIndex - 1, ArrayType, col
        quickSort arr, partitionIndex + 1, high, ArrayType, col
    End If
End Sub

Private Function qsPartition(ByVal arr As Variant, ByVal low As Long, ByVal high As Long, ByVal ArrayType As ArrayTypes, Optional ByVal col As Long) As Long
    Dim i As Long, j As Long
    Dim pivot As Variant, current As Variant
    
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
            qsSwap arr, i, j
        End If
    Next
    qsSwap arr, i + 1, high
    qsPartition = i + 1
End Function

Private Sub qsSwap(ByVal arr As Variant, ByVal i As Long, ByVal j As Long)
    Dim temporaryArray As Variant
    temporaryArray = arr(i)
    arr(i) = arr(j)
    arr(j) = temporaryArray
End Sub

Private Function getInternalItems() As Variant
    Dim result() As Variant
    result = pItems
    If Me.Capacity > pLength Then
        If pLength > 0 Then ReDim Preserve result(pBase To Me.UpperBound)
    End If
    getInternalItems = result
End Function

Private Function recursiveToString(ByVal arr As Variant, ByVal PrettyPrint As Boolean, Optional ByVal tabs As Long = 1) As String
    Const tabWidth As Long = 2
    Dim result As String, sep As String
    Dim i As Long
    result = "["
    For i = LBound(arr) To UBound(arr)
        sep = IIf(i = UBound(arr), "]", ", ")
        If IsArray(arr(i)) Then
            If PrettyPrint Then result = result & vbCrLf & Space(tabWidth * tabs)
            result = result & recursiveToString(arr(i), PrettyPrint, tabs + 1)
            result = result & IIf(i = UBound(arr) And PrettyPrint, vbCrLf & Space(tabWidth * (tabs - 1)) & sep, sep)
        ElseIf IsObject(arr(i)) Then
            result = result & "OBJECT" & sep
        Else
            result = result & Replace(CStr(arr(i)), ",", vbNullString) & sep
        End If
    Next
    recursiveToString = result
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

Private Function Rebase(Optional ByRef recurse As Boolean, Optional ByVal arr As Variant) As Variant
    Dim isMD As Boolean, applyToMe As Boolean
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

Private Function recursiveRebase(ByVal arr As Variant, ByVal recurse As Boolean) As Variant
    Dim i As Long, localBase As Long, localUpperBound As Long, offset As Long
    Dim newItems() As Variant
    
    localBase = LBound(arr)
    localUpperBound = UBound(arr)
    offset = pBase - localBase
    ReDim newItems(pBase To localUpperBound + offset)
    For i = localBase To localUpperBound
        If IsArray(arr(i)) And recurse Then
            newItems(i + offset) = recursiveRebase(arr(i), recurse)
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

Private Function copyRange(ByRef newItems As Variant) As Variant
    Dim i As Long
    For i = LBound(newItems) To UBound(newItems)
        newItems(i) = pItems(i)
    Next
    copyRange = newItems
End Function

Private Function jaggedToMulti(ByVal arr As Variant, _
                               Optional ByVal depth As Long, _
                               Optional ByVal currentDepth As Long, _
                               Optional ByRef crumbs As Variant, _
                               Optional ByRef result As Variant) As Variant
    Dim localBase As Long
    Dim localUpperBound As Long
    Dim localDepth As Long
    Dim localCurrentDepth As Long
    Dim i As Long
    
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
        ' handle error
        Debug.Print "Cannot convert arrays with more than 20 dimensions to a multidimensional array!"
    ElseIf localDepth > 1 Then
        If Not IsArray(crumbs) Then ReDim crumbs(localDepth - 1)
        If Not IsArray(result) Then
            result = createMultidimensionalArray(mapJaggedArray(arr))
        End If
        localCurrentDepth = localCurrentDepth + 1
        localBase = LBound(arr)
        localUpperBound = UBound(arr)
        For i = localBase To localUpperBound
            crumbs(localCurrentDepth - 1) = i
            If localCurrentDepth = localDepth Or Not IsArray(arr(i)) Then
                result = letElementByBreadcrumb(result, crumbs, arr(i))
            Else
                result = jaggedToMulti(arr(i), localDepth, localCurrentDepth, crumbs, result)
            End If
        Next
    Else
        result = arr
    End If
    
    jaggedToMulti = result
End Function

Private Function multiToJagged(ByVal arr As Variant, Optional ByVal depth As Long, Optional ByVal currentDepth As Long, Optional ByRef crumbs As Variant) As Variant
    Dim localBase As Long
    Dim localUpperBound As Long
    Dim localDepth As Long
    Dim localCurrentDepth As Long
    Dim i As Long
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
        If Not IsArray(crumbs) Then ReDim crumbs(localDepth - 1)
        localCurrentDepth = localCurrentDepth + 1
        localBase = LBound(arr, localCurrentDepth)
        localUpperBound = UBound(arr, localCurrentDepth)
        ReDim result(localBase To localUpperBound)
        For i = localBase To localUpperBound
            crumbs(localCurrentDepth - 1) = i
            If localCurrentDepth = localDepth Then
                result(i) = getElementByBreadcrumb(arr, crumbs)
            Else
                result(i) = multiToJagged(arr, localDepth, localCurrentDepth, crumbs)
            End If
        Next
    Else
        result = arr
    End If
    multiToJagged = result
End Function

Private Function getJaggedArrayDepth(ByVal arr As Variant) As Variant
    Dim i As Long, localBase As Long, localUpperBound As Long, currentDepth As Long, maxDepth As Long
    Dim depth As Long
    If IsArray(arr) Then
        depth = depth + 1
        localBase = LBound(arr)
        localUpperBound = UBound(arr)
        For i = localBase To localUpperBound
            If IsArray(arr(i)) Then
                currentDepth = getJaggedArrayDepth(arr(i))
            End If
            If currentDepth > maxDepth Then maxDepth = currentDepth
        Next
        depth = depth + maxDepth
    End If
    getJaggedArrayDepth = depth
End Function

Private Function mapJaggedArray(ByVal arr As Variant, Optional ByRef map As Variant, Optional ByVal currentIndex As Long) As Variant
    Dim i As Long
    Dim bounds() As Variant
    Dim Element As Variant
    
    If Not IsArray(arr) Then Exit Function
    If Not IsArray(map) Then
        ReDim map(0 To getJaggedArrayDepth(arr) - 1)
        For i = LBound(map) To UBound(map)
            map(i) = Array(0, 0)
        Next
    End If
    
    bounds = getArrayBounds(arr)
    
    If bounds(0) < map(currentIndex)(0) Then map(currentIndex)(0) = bounds(0)
    If bounds(1) > map(currentIndex)(1) Then map(currentIndex)(1) = bounds(1)
    
    For Each Element In arr
        If IsArray(Element) Then mapJaggedArray Element, map, currentIndex + 1
    Next
    
    mapJaggedArray = map
End Function

Private Function getArrayBounds(ByVal arr As Variant) As Variant
    Dim result(0 To 1) As Variant
    If IsArray(arr) Then
        result(0) = LBound(arr)
        result(1) = UBound(arr)
    End If
    getArrayBounds = result
End Function

Private Function getMultidimensionalArrayDepth(ByVal arr As Variant) As Long
    '@Ignore VariableNotUsed
    Dim i As Long, void As Long
    On Error Resume Next
    Do
        i = i + 1
        void = UBound(arr, i)
    Loop Until Err.Number <> 0
    Err.Clear
    On Error GoTo 0
    getMultidimensionalArrayDepth = i - 1
End Function

Private Function getElementByBreadcrumb(ByVal arr As Variant, ByVal crumb As Variant) As Variant
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

Private Function letElementByBreadcrumb(ByVal arr As Variant, ByVal crumb As Variant, ByVal Element As Variant) As Variant
    If IsObject(Element) Then
        Select Case UBound(crumb)
        Case 0
            Set arr(crumb(0)) = Element
        Case 1
            Set arr(crumb(0), crumb(1)) = Element
        Case 2
            Set arr(crumb(0), crumb(1), crumb(2)) = Element
        Case 3
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3)) = Element
        Case 4
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4)) = Element
        Case 5
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5)) = Element
        Case 6
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6)) = Element
        Case 7
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7)) = Element
        Case 8
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8)) = Element
        Case 9
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9)) = Element
        Case 10
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10)) = Element
        Case 11
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                    crumb(11)) = Element
        Case 12
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                    crumb(11), crumb(12)) = Element
        Case 13
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                    crumb(11), crumb(12), crumb(13)) = Element
        Case 14
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                    crumb(11), crumb(12), crumb(13), crumb(14)) = Element
        Case 15
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                    crumb(11), crumb(12), crumb(13), crumb(14), crumb(15)) = Element
        Case 16
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                    crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16)) = Element
        Case 17
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                    crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17)) = Element
        Case 18
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                    crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18)) = Element
        Case 19
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                    crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18), crumb(19)) = Element
        Case 20
            Set arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                    crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18), crumb(19), crumb(20)) = Element
        End Select
    Else
        Select Case UBound(crumb)
        Case 0
            arr(crumb(0)) = Element
        Case 1
            arr(crumb(0), crumb(1)) = Element
        Case 2
            arr(crumb(0), crumb(1), crumb(2)) = Element
        Case 3
            arr(crumb(0), crumb(1), crumb(2), crumb(3)) = Element
        Case 4
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4)) = Element
        Case 5
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5)) = Element
        Case 6
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6)) = Element
        Case 7
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7)) = Element
        Case 8
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8)) = Element
        Case 9
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9)) = Element
        Case 10
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10)) = Element
        Case 11
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                crumb(11)) = Element
        Case 12
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                crumb(11), crumb(12)) = Element
        Case 13
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                crumb(11), crumb(12), crumb(13)) = Element
        Case 14
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                crumb(11), crumb(12), crumb(13), crumb(14)) = Element
        Case 15
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                crumb(11), crumb(12), crumb(13), crumb(14), crumb(15)) = Element
        Case 16
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16)) = Element
        Case 17
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17)) = Element
        Case 18
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18)) = Element
        Case 19
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18), crumb(19)) = Element
        Case 20
            arr(crumb(0), crumb(1), crumb(2), crumb(3), crumb(4), crumb(5), crumb(6), crumb(7), crumb(8), crumb(9), crumb(10), _
                crumb(11), crumb(12), crumb(13), crumb(14), crumb(15), crumb(16), crumb(17), crumb(18), crumb(19), crumb(20)) = Element
        End Select
    End If
    letElementByBreadcrumb = arr
End Function

Private Function createMultidimensionalArray(ByVal crumb As Variant) As Variant
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




