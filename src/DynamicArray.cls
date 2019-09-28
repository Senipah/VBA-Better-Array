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
    jagged
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
Private Sub Class_Terminate()

End Sub

'''''''''''''''''''''
' Public Properties '
'''''''''''''''''''''

Public Property Let Capacity(value As Long)
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

' Read Only
'Public Property Let Length(value As Long)
'    pLength = value
'End Property

Public Property Get length() As Long
    length = pLength
End Property

Public Property Get Items()
    Dim rtn() As Variant
    rtn = getInternalItems()
    If pArrayType = ArrayTypes.MultiDimension And _
            getArrayType(rtn) = ArrayTypes.jagged Then rtn = jaggedToMulti(rtn)
    Items = rtn
End Property

Public Property Let Items(values As Variant)
    Dim lb As Long, ub As Long, i As Long
    
    pArrayType = getArrayType(values)
    If pArrayType = ArrayTypes.Undefined _
                Or pArrayType = ArrayTypes.Unallocated Then
        Err.Raise 9
    Else
        If pArrayType = ArrayTypes.MultiDimension Then
            values = multiToJagged(values)
        End If
        lb = LBound(values)
        pItems = values
        pLength = UBound(pItems) - LBound(pItems) + 1
        If lb <> pBase Then
            Rebase
        End If
        Me.Capacity = pLength
    End If
End Property

Public Property Get UpperBound()
    UpperBound = pLength + pBase - 1
End Property

Public Property Get Base() As Long
    Base = pBase
End Property

Public Property Let Base(value As Long)
    If value <> pBase And value >= 0 Then
        pBase = value
        Rebase
    End If
End Property

Public Property Let Item(index As Long, element As Variant)
    If pLength >= index Then
        If IsObject(element) Then
            Set pItems(index) = element
        Else
            pItems(index) = element
        End If
    Else
        Me.Push element
    End If
End Property

Public Property Get Item(index As Long)
    If index <= pLength Then
        If IsObject(pItems(index)) Then
            Set Item = pItems(index)
        Else
            Item = pItems(index)
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

Public Function Push(ParamArray args() As Variant) As Variant
    Dim e As Variant
    If pArrayType = ArrayTypes.Unallocated Or _
            pArrayType = ArrayTypes.Undefined Then
        pArrayType = ArrayTypes.OneDimension
    End If
    For Each e In args
        If IsArray(e) And pArrayType = ArrayTypes.OneDimension Then
            pArrayType = ArrayTypes.jagged
        End If
        If pLength = pCapacity Then
            EnsureCapacity pLength + 1
        End If
        If IsObject(e) Then
            Set pItems(pLength) = e
        Else
            pItems(pLength) = e
        End If
        pLength = pLength + 1
    Next
    Push = pLength
End Function

Public Function Pop() As Variant
    Dim newItems() As Variant
    Dim rtn As Variant
    Dim i As Long
    rtn = pItems(Me.UpperBound)
    pLength = pLength - 1
    ReDim newItems(pBase To Me.UpperBound)
    Me.Items = copyRange(newItems)
    Pop = rtn
End Function

Public Function Shift() As Variant
    Dim rtn As Variant
    Dim newItems() As Variant
    rtn = pItems(pBase)
    ReDim newItems((pBase + 1) To Me.UpperBound)
    Me.Items = copyRange(newItems)
    Shift = rtn
End Function

Public Function Unshift(ParamArray args() As Variant) As Long
    Dim newItems() As Variant, oldItems() As Variant
    newItems = args
    oldItems = pItems
    Me.Items = newItems
    Me.Concat oldItems
    Unshift = Me.length
End Function

Public Function Concat(ParamArray args() As Variant) As Variant
    Dim arg As Variant, e As Variant
    For Each arg In args
        If IsArray(arg) Then
            For Each e In arg
                Me.Push e
            Next
        Else
            Me.Push arg
        End If
    Next
    Concat = Me.Items
End Function

Public Function CopyFromCollection(c As Collection) As Variant
    If IsMissing(c) Then Err.Raise 448
    Dim i As Long
    Dim newItems() As Variant
    pLength = c.count
    If pLength = 0 Then
        newItems = pEmptyArray
    Else
        ReDim newItems(pBase To (pLength - pBase - 1))
        For i = 1 To pLength
            newItems(i + pBase - 1) = c.Item(i)
        Next
    End If
    Me.Items = newItems
    CopyFromCollection = newItems
End Function

Public Function ToString(Optional prettyPrint As Boolean) As String
    Dim arrType As ArrayTypes
    Dim rtn As String
    Dim Items As Variant
    
    Items = getInternalItems()
    arrType = getArrayType(Items)
    If Not arrType = ArrayTypes.Undefined And _
            Not arrType = ArrayTypes.Unallocated Then
        If arrType = ArrayTypes.MultiDimension Then
            Items = multiToJagged(Items)
        End If
        rtn = recursiveToString(Items, prettyPrint)
    End If
    ToString = rtn
End Function

Public Function Sort(Optional col As Long) As Variant
    Dim Items As Variant, sortedItems As Variant
    Dim arrType As ArrayTypes
    
    Items = getInternalItems()
    arrType = getArrayType(Items)
    If arrType = ArrayTypes.Unallocated Or _
            arrType = ArrayTypes.Undefined Then
        sortedItems = pEmptyArray
    Else
        If Me.length > 0 Then
            If arrType <> OneDimension Then
                If arrType = ArrayTypes.MultiDimension Then
                    Items = multiToJagged(Items)
                End If
                Dim d As Long
                d = getJaggedArrayDepth(Items)
                If d > 2 Then
                    Debug.Print "Cannot sort greater than 2 dimensions"
                    Exit Function
                End If
            End If
            quickSort Items, LBound(Items), UBound(Items), arrType, col
        End If
        sortedItems = Items
    End If
    Me.Items = sortedItems
    Sort = sortedItems
End Function

Public Function CopyWithin(target As Long, Optional startI As Long, Optional endI As Long) As Variant
    Dim lngth As Long, relativeTarget As Long, relativeStart As Long, relativeEnd As Long
    Dim idx As Long, from As Long, final As Long, count As Long, direction As Long
    Dim arr As Variant
    
    arr = getInternalItems()
    lngth = Me.length
    relativeTarget = target
    
    If relativeTarget < 0 Then
        idx = Max((lngth + relativeTarget), 0)
    Else
        idx = Min(relativeTarget, lngth)
    End If
    
    relativeStart = startI
    
    If relativeStart < 0 Then
        from = Max((lngth + relativeStart), 0)
    Else
        from = Min(relativeStart, lngth)
    End If
    
    If endI = 0 Then
        relativeEnd = lngth
    Else
        relativeEnd = endI
    End If
    
    If relativeEnd < 0 Then
        final = Max((lngth + relativeEnd), 0)
    Else
        final = Min(relativeEnd, lngth)
    End If
    
    count = Min(final - from, lngth - idx)
    
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
Public Function Filter(match As Variant, Optional exclude As Boolean) As Variant
    Dim v As Variant
    Dim oldArr() As Variant, newArr() As Variant
    oldArr = getInternalItems()
    ReDim newArr(LBound(oldArr), UBound(oldArr))
    Dim i As Long
    i = Me.Base
    For Each v In oldArr
        If (exclude And v = match) Or (Not exclude And v <> match) Then
            newArr(i) = v
            i = i + 1
        End If
    Next
    ReDim Preserve newArr(Me.Base, i)
    Me.Items = newArr
    Filter = newArr
End Function

Public Function Includes(searchElement As Variant, Optional fromIndex As Long) As Boolean
    Dim arr As Variant
    arr = getInternalItems()
    Dim lngth As Long, k As Long
    lngth = Me.length
    If lngth = 0 Then
        Includes = False
        Exit Function
    End If
    If IsMissing(fromIndex) Then fromIndex = Me.Base
    
    If fromIndex >= Me.Base Then
        k = fromIndex
    Else
        k = lngth + fromIndex
    End If
    
    While k < lngth
        If arr(k) = searchElement Then
            Includes = True
            Exit Function
        End If
        k = k + 1
    Wend
    Includes = False
End Function

Public Function Keys() As Variant
    Dim i As Long, lb As Long, ub As Long
    lb = Me.Base
    ub = Me.UpperBound
    Dim rtn() As Variant: ReDim rtn(lb To ub)
    For i = lb To ub: rtn(i) = i: Next
    Keys = rtn
End Function

'TODO: finish & write docs
Private Function Max(ParamArray args() As Variant) As Variant
    Dim i As Long
    Dim arr() As Variant
    Dim rtn As Variant, v As Variant
    
    If IsMissing(args) Then
        arr = getInternalItems()
    Else
        arr = args
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
            If IsEmpty(rtn) Then
                rtn = arr(i)
            Else
                If arr(i) > rtn Then rtn = arr(i)
            End If
        End If
    Next
    
    Max = rtn
End Function

'TODO: finish & write docs
Private Function Min(ParamArray args() As Variant) As Variant
    Dim i As Long
    Dim arr() As Variant
    Dim rtn As Variant, v As Variant
    
    If IsMissing(args) Then
        arr = getInternalItems()
    Else
        arr = args
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
            If IsEmpty(rtn) Then
                rtn = arr(i)
            Else
                If arr(i) > rtn Then rtn = arr(i)
            End If
        End If
    Next
    
    Min = rtn
End Function

'TODO: finish & write docs
Public Function Slice(startI As Long, Optional endI As Long) As Variant
    '@Ignore ShadowedDeclaration
    Dim lngth As Long, relativeStart As Long, relativeEnd As Long
    Dim k As Long, final As Long, count As Long, n As Long
    Dim arr() As Variant, a() As Variant
    arr = getInternalItems()
    lngth = Me.length
    relativeStart = startI
    If relativeStart < 0 Then
        k = Max((lngth + relativeStart), 0)
    Else
        k = Min(relativeStart, lngth)
    End If
    If endI = 0 Then
        relativeEnd = lngth
    Else
        relativeEnd = endI
    End If
    If relativeEnd < 0 Then
        final = Max((lngth + relativeEnd), 0)
    Else
        final = Min(relativeEnd, lngth)
    End If
    
    count = Max(final - k, 0)
    ReDim a(n To count - 1)
    Do While k < final
        If k >= LBound(arr) And k <= UBound(arr) Then
            a(n) = arr(k)
            n = n + 1
            k = k + 1
        End If
    Loop
    
    Slice = a
End Function

'TODO: finish & write docs
Public Function Reverse() As Variant
    Dim i As Long, ub As Long
    Dim arr() As Variant, rtn() As Variant
    arr = getInternalItems()
    ub = Me.UpperBound
    ReDim rtn(Me.Base To ub)
    For i = ub To Me.Base Step -1
        rtn(ub - i) = arr(i)
    Next
    Me.Items = rtn
    Reverse = rtn
End Function

'''''''''''''''''''
' Private Methods '
'''''''''''''''''''

Private Sub quickSort(ByRef arr As Variant, low As Long, high As Long, arrType As ArrayTypes, Optional col As Long)
    Dim pi As Long                               ' Partition index
    If low < high Then
        pi = qsPartition(arr, low, high, arrType, col)
        quickSort arr, low, pi - 1, arrType, col
        quickSort arr, pi + 1, high, arrType, col
    End If
End Sub

Private Function qsPartition(ByRef arr As Variant, low As Long, high As Long, arrType As ArrayTypes, Optional col As Long) As Long
    Dim i As Long, j As Long
    Dim pivot As Variant, current As Variant
    
    If arrType = jagged Then
        pivot = arr(high)(col)
    Else
        pivot = arr(high)
    End If
    
    i = low - 1
    For j = low To high - 1
        If arrType = jagged Then
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

Private Sub qsSwap(arr As Variant, i As Long, j As Long)
    Dim tmp As Variant
    tmp = arr(i)
    arr(i) = arr(j)
    arr(j) = tmp
End Sub

Private Function getInternalItems() As Variant
    Dim rtn() As Variant
    rtn = pItems
    If Me.Capacity > pLength Then
        If pLength > 0 Then ReDim Preserve rtn(pBase To Me.UpperBound)
    End If
    getInternalItems = rtn
End Function

Private Function recursiveToString(arr As Variant, prettyPrint As Boolean, Optional ByVal tabs As Long = 1) As String
    Const tabWidth As Long = 2
    Dim rtn As String, sep As String
    Dim last As Boolean
    Dim i As Long
    rtn = "["
    For i = LBound(arr) To UBound(arr)
        sep = IIf(i = UBound(arr), "]", ", ")
        If IsArray(arr(i)) Then
            If prettyPrint Then rtn = rtn & vbCrLf & Space(tabWidth * tabs)
            rtn = rtn & recursiveToString(arr(i), prettyPrint, tabs + 1)
            rtn = rtn & IIf(i = UBound(arr) And prettyPrint, vbCrLf & Space(tabWidth * (tabs - 1)) & sep, sep)
        ElseIf IsObject(arr(i)) Then
            rtn = rtn & "OBJECT" & sep
        Else
            rtn = rtn & Replace(CStr(arr(i)), ",", vbNullString) & sep
        End If
    Next
    recursiveToString = rtn
End Function

Private Function getArrayType(arr As Variant) As ArrayTypes
    Dim rtn As ArrayTypes
    If Not isArrayAllocated(arr) Then
        rtn = ArrayTypes.Unallocated
    Else
        If isMultidimensionalArray(arr) Then
            rtn = ArrayTypes.MultiDimension
        ElseIf isJaggedArray(arr) Then
            rtn = ArrayTypes.jagged
        Else
            rtn = ArrayTypes.OneDimension
        End If
    End If
    getArrayType = rtn
End Function

Private Function isArrayAllocated(arr As Variant) As Boolean
        On Error Resume Next
        isArrayAllocated = IsArray(arr) And _
                           Not IsError(LBound(arr, 1)) And _
                           LBound(arr, 1) <= UBound(arr, 1)
End Function

Private Function isJaggedArray(arr As Variant) As Boolean
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

Private Function isMultidimensionalArray(arr As Variant) As Boolean
    If IsArray(arr) Then
        On Error GoTo errHandler
        Dim ub As Long
        ub = UBound(arr, 2)
        isMultidimensionalArray = True
        On Error GoTo 0
    End If
    Exit Function
errHandler:
    Err.Clear
End Function

Private Function Rebase(Optional recurse As Boolean, Optional arr As Variant) As Variant
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

Private Function recursiveRebase(arr As Variant, recurse As Boolean)
    Dim i As Long, lb As Long, ub As Long, offset As Long
    Dim newItems() As Variant
    
    lb = LBound(arr)
    ub = UBound(arr)
    offset = pBase - lb
    ReDim newItems(pBase To ub + offset)
    For i = lb To ub
        If IsArray(arr(i)) And recurse Then
            newItems(i + offset) = recursiveRebase(arr(i), recurse)
        Else
            newItems(i + offset) = arr(i)
        End If
    Next
    
    recursiveRebase = newItems
End Function

Private Sub EnsureCapacity(Min As Long)
    If Me.Capacity < Min Then
        Dim newCapacity As Long
        newCapacity = IIf(Me.Capacity = 0, DEFAULT_CAPACITY, Me.Capacity * 2)
        If newCapacity > MAX_ARRAY_LENGTH Then newCapacity = MAX_ARRAY_LENGTH
        If newCapacity < Min Then newCapacity = Min
        Me.Capacity = newCapacity
    End If
End Sub

Private Function copyRange(newItems As Variant) As Variant
    Dim i As Long
    For i = LBound(newItems) To UBound(newItems)
        newItems(i) = pItems(i)
    Next
    copyRange = newItems
End Function

Private Function jaggedToMulti(arr As Variant, _
            Optional depth As Long, _
            Optional ByVal d As Long, _
            Optional crumbs As Variant, _
            Optional rtn As Variant) As Variant
    Dim lb As Long, ub As Long, i As Long
       
    If depth = 0 Then
        If Not isJaggedArray(arr) Then
            Debug.Print "Argument passed to 'jaggedToMulti' not a jagged array"
            Exit Function
        End If
        depth = getJaggedArrayDepth(arr)
    End If
    
    If depth > 20 Then
        ' too many dimensions to be handled by getElementByBreadcrumb
        ' handle error
        Debug.Print "Cannot convert arrays with more than 20 dimensions to a multidimensional array!"
    ElseIf depth > 1 Then
        If Not IsArray(crumbs) Then ReDim crumbs(depth - 1)
        If Not IsArray(rtn) Then
            rtn = createMultidimensionalArray(mapJaggedArray(arr))
        End If
        d = d + 1
        lb = LBound(arr)
        ub = UBound(arr)
        For i = lb To ub
            crumbs(d - 1) = i
            If d = depth Or Not IsArray(arr(i)) Then
                rtn = letElementByBreadcrumb(rtn, crumbs, arr(i))
            Else
                rtn = jaggedToMulti(arr(i), depth, d, crumbs, rtn)
            End If
        Next
    Else
        rtn = arr
    End If
    
    jaggedToMulti = rtn
End Function

Private Function multiToJagged(arr As Variant, Optional depth As Long, Optional ByVal d As Long, Optional crumbs As Variant) As Variant
    Dim lb As Long, ub As Long, i As Long
    Dim rtn() As Variant
    
    If depth = 0 Then
        If Not isMultidimensionalArray(arr) Then
            Debug.Print "Argument passed to 'multiToJagged' not a multidimensional array"
            Exit Function
        End If
        depth = getMultidimensionalArrayDepth(arr)
    End If
    
    If depth > 20 Then
        ' too many dimensions to be handled by getElementByBreadcrumb
        ' handle error
        Debug.Print "Cannot convert arrays with more than 20 dimensions to a jagged array!"
    ElseIf depth > 1 Then
        If Not IsArray(crumbs) Then ReDim crumbs(depth - 1)
        d = d + 1
        lb = LBound(arr, d)
        ub = UBound(arr, d)
        ReDim rtn(lb To ub)
        For i = lb To ub
            crumbs(d - 1) = i
            If d = depth Then
                rtn(i) = getElementByBreadcrumb(arr, crumbs)
            Else
                rtn(i) = multiToJagged(arr, depth, d, crumbs)
            End If
        Next
    Else
        rtn = arr
    End If
    multiToJagged = rtn
End Function

Private Function getJaggedArrayDepth(arr As Variant) As Variant
    Dim i As Long, lb As Long, ub As Long, x As Long, y As Long
    Dim depth As Long
    If IsArray(arr) Then
        depth = depth + 1
        lb = LBound(arr)
        ub = UBound(arr)
        For i = lb To ub
            If IsArray(arr(i)) Then
                x = getJaggedArrayDepth(arr(i))
            End If
            If x > y Then y = x
        Next
        depth = depth + y
    End If
    getJaggedArrayDepth = depth
End Function

Private Function mapJaggedArray(arr As Variant, Optional map As Variant, Optional loc As Long) As Variant
    Dim i As Long
    Dim x() As Variant
    Dim element As Variant
    
    If Not IsArray(arr) Then Exit Function
    If Not IsArray(map) Then
        ReDim map(0 To getJaggedArrayDepth(arr) - 1)
        For i = LBound(map) To UBound(map)
            map(i) = Array(0, 0)
        Next
    End If
    
    x = getArrayBounds(arr)
    
    If x(0) < map(loc)(0) Then map(loc)(0) = x(0)
    If x(1) > map(loc)(1) Then map(loc)(1) = x(1)
    
    For Each element In arr
        If IsArray(element) Then mapJaggedArray element, map, loc + 1
    Next
    
    mapJaggedArray = map
End Function

Private Function getArrayBounds(arr As Variant) As Variant
    Dim rtn(0 To 1) As Variant
    If IsArray(arr) Then
        rtn(0) = LBound(arr)
        rtn(1) = UBound(arr)
    End If
    getArrayBounds = rtn
End Function

Private Function getMultidimensionalArrayDepth(arr As Variant) As Long
    Dim i As Long, void As Long
    On Error Resume Next
    Do
        i = i + 1
        void = UBound(arr, i)
    Loop Until Err.Number <> 0
    Err.Clear
    getMultidimensionalArrayDepth = i - 1
End Function


Private Function getElementByBreadcrumb(arr As Variant, c As Variant) As Variant
    Dim rtn As Variant
    Select Case UBound(c)
        Case 0
            rtn = arr(c(0))
        Case 1
            rtn = arr(c(0), c(1))
        Case 2
            rtn = arr(c(0), c(1), c(2))
        Case 3
            rtn = arr(c(0), c(1), c(2), c(3))
        Case 4
            rtn = arr(c(0), c(1), c(2), c(3), c(4))
        Case 5
            rtn = arr(c(0), c(1), c(2), c(3), c(4), c(5))
        Case 6
            rtn = arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6))
        Case 7
            rtn = arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7))
        Case 8
            rtn = arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8))
        Case 9
            rtn = arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9))
        Case 10
            rtn = arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10))
        Case 11
            rtn = arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), _
                    c(11))
        Case 12
            rtn = arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), _
                    c(11), c(12))
        Case 13
            rtn = arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), _
                    c(11), c(12), c(13))
        Case 14
            rtn = arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), _
                    c(11), c(12), c(13), c(14))
        Case 15
            rtn = arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), _
                    c(11), c(12), c(13), c(14), c(15))
        Case 16
            rtn = arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), _
                    c(11), c(12), c(13), c(14), c(15), c(16))
        Case 17
            rtn = arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), _
                    c(11), c(12), c(13), c(14), c(15), c(16), c(17))
        Case 18
            rtn = arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), _
                    c(11), c(12), c(13), c(14), c(15), c(16), c(17), c(18))
        Case 19
            rtn = arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), _
                    c(11), c(12), c(13), c(14), c(15), c(16), c(17), c(18), c(19))
        Case 20
            rtn = arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), _
                    c(11), c(12), c(13), c(14), c(15), c(16), c(17), c(18), c(19), c(20))
    End Select
    getElementByBreadcrumb = rtn
End Function

Private Function letElementByBreadcrumb(arr As Variant, c As Variant, e As Variant) As Variant
    If IsObject(e) Then
        Select Case UBound(c)
            Case 0
                Set arr(c(0)) = e
            Case 1
                Set arr(c(0), c(1)) = e
            Case 2
                Set arr(c(0), c(1), c(2)) = e
            Case 3
                Set arr(c(0), c(1), c(2), c(3)) = e
            Case 4
                Set arr(c(0), c(1), c(2), c(3), c(4)) = e
            Case 5
                Set arr(c(0), c(1), c(2), c(3), c(4), c(5)) = e
            Case 6
                Set arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6)) = e
            Case 7
                Set arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7)) = e
            Case 8
                Set arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8)) = e
            Case 9
                Set arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9)) = e
            Case 10
                Set arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10)) = e
            Case 11
                Set arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), _
                        c(11)) = e
            Case 12
                Set arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), _
                        c(11), c(12)) = e
            Case 13
                Set arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), _
                        c(11), c(12), c(13)) = e
            Case 14
                Set arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), _
                        c(11), c(12), c(13), c(14)) = e
            Case 15
                Set arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), _
                        c(11), c(12), c(13), c(14), c(15)) = e
            Case 16
                Set arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), _
                        c(11), c(12), c(13), c(14), c(15), c(16)) = e
            Case 17
                Set arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), _
                        c(11), c(12), c(13), c(14), c(15), c(16), c(17)) = e
            Case 18
                Set arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), _
                        c(11), c(12), c(13), c(14), c(15), c(16), c(17), c(18)) = e
            Case 19
                Set arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), _
                        c(11), c(12), c(13), c(14), c(15), c(16), c(17), c(18), c(19)) = e
            Case 20
                Set arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), _
                        c(11), c(12), c(13), c(14), c(15), c(16), c(17), c(18), c(19), c(20)) = e
        End Select
    Else
        Select Case UBound(c)
            Case 0
                arr(c(0)) = e
            Case 1
                arr(c(0), c(1)) = e
            Case 2
                arr(c(0), c(1), c(2)) = e
            Case 3
                arr(c(0), c(1), c(2), c(3)) = e
            Case 4
                arr(c(0), c(1), c(2), c(3), c(4)) = e
            Case 5
                arr(c(0), c(1), c(2), c(3), c(4), c(5)) = e
            Case 6
                arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6)) = e
            Case 7
                arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7)) = e
            Case 8
                arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8)) = e
            Case 9
                arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9)) = e
            Case 10
                arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10)) = e
            Case 11
                arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), _
                        c(11)) = e
            Case 12
                arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), _
                        c(11), c(12)) = e
            Case 13
                arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), _
                        c(11), c(12), c(13)) = e
            Case 14
                arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), _
                        c(11), c(12), c(13), c(14)) = e
            Case 15
                arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), _
                        c(11), c(12), c(13), c(14), c(15)) = e
            Case 16
                arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), _
                        c(11), c(12), c(13), c(14), c(15), c(16)) = e
            Case 17
                arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), _
                        c(11), c(12), c(13), c(14), c(15), c(16), c(17)) = e
            Case 18
                arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), _
                        c(11), c(12), c(13), c(14), c(15), c(16), c(17), c(18)) = e
            Case 19
                arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), _
                        c(11), c(12), c(13), c(14), c(15), c(16), c(17), c(18), c(19)) = e
            Case 20
                arr(c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), _
                        c(11), c(12), c(13), c(14), c(15), c(16), c(17), c(18), c(19), c(20)) = e
        End Select
    End If
    letElementByBreadcrumb = arr
End Function

Private Function createMultidimensionalArray(c As Variant) As Variant
    Dim rtn() As Variant
    Select Case UBound(c)
        Case 0
            ReDim rtn(c(0)(0) To c(0)(1))
        Case 1
            ReDim rtn(c(0)(0) To c(0)(1), c(1)(0) To c(1)(1))
        Case 2
            ReDim rtn(c(0)(0) To c(0)(1), c(1)(0) To c(1)(1), c(2)(0) To c(2)(1))
        Case 3
            ReDim rtn(c(0)(0) To c(0)(1), c(1)(0) To c(1)(1), c(2)(0) To c(2)(1), c(3)(0) To c(3)(1))
        Case 4
            ReDim rtn(c(0)(0) To c(0)(1), c(1)(0) To c(1)(1), c(2)(0) To c(2)(1), c(3)(0) To c(3)(1), _
                c(4)(0) To c(4)(1))
        Case 5
            ReDim rtn(c(0)(0) To c(0)(1), c(1)(0) To c(1)(1), c(2)(0) To c(2)(1), c(3)(0) To c(3)(1), _
                c(4)(0) To c(4)(1), c(5)(0) To c(5)(1))
        Case 6
        Case 5
            ReDim rtn(c(0)(0) To c(0)(1), c(1)(0) To c(1)(1), c(2)(0) To c(2)(1), c(3)(0) To c(3)(1), _
                c(4)(0) To c(4)(1), c(5)(0) To c(5)(1), c(6)(0) To c(6)(1))
        Case 7
            ReDim rtn(c(0)(0) To c(0)(1), c(1)(0) To c(1)(1), c(2)(0) To c(2)(1), c(3)(0) To c(3)(1), _
                c(4)(0) To c(4)(1), c(5)(0) To c(5)(1), c(6)(0) To c(6)(1), c(7)(0) To c(7)(1))
        Case 8
            ReDim rtn(c(0)(0) To c(0)(1), c(1)(0) To c(1)(1), c(2)(0) To c(2)(1), c(3)(0) To c(3)(1), _
                c(4)(0) To c(4)(1), c(5)(0) To c(5)(1), c(6)(0) To c(6)(1), c(7)(0) To c(7)(1), _
                c(8)(0) To c(8)(1))
        Case 9
            ReDim rtn(c(0)(0) To c(0)(1), c(1)(0) To c(1)(1), c(2)(0) To c(2)(1), c(3)(0) To c(3)(1), _
                c(4)(0) To c(4)(1), c(5)(0) To c(5)(1), c(6)(0) To c(6)(1), c(7)(0) To c(7)(1), _
                c(8)(0) To c(8)(1), c(9)(0) To c(9)(1))
        Case 10
            ReDim rtn(c(0)(0) To c(0)(1), c(1)(0) To c(1)(1), c(2)(0) To c(2)(1), c(3)(0) To c(3)(1), _
                c(4)(0) To c(4)(1), c(5)(0) To c(5)(1), c(6)(0) To c(6)(1), c(7)(0) To c(7)(1), _
                c(8)(0) To c(8)(1), c(9)(0) To c(9)(1), c(10)(0) To c(10)(1))
        Case 11
            ReDim rtn(c(0)(0) To c(0)(1), c(1)(0) To c(1)(1), c(2)(0) To c(2)(1), c(3)(0) To c(3)(1), _
                c(4)(0) To c(4)(1), c(5)(0) To c(5)(1), c(6)(0) To c(6)(1), c(7)(0) To c(7)(1), _
                c(8)(0) To c(8)(1), c(9)(0) To c(9)(1), c(10)(0) To c(10)(1), c(11)(0) To c(11)(1))
        Case 12
            ReDim rtn(c(0)(0) To c(0)(1), c(1)(0) To c(1)(1), c(2)(0) To c(2)(1), c(3)(0) To c(3)(1), _
                c(4)(0) To c(4)(1), c(5)(0) To c(5)(1), c(6)(0) To c(6)(1), c(7)(0) To c(7)(1), _
                c(8)(0) To c(8)(1), c(9)(0) To c(9)(1), c(10)(0) To c(10)(1), c(11)(0) To c(11)(1), _
                c(12)(0) To c(12)(1))
        Case 13
            ReDim rtn(c(0)(0) To c(0)(1), c(1)(0) To c(1)(1), c(2)(0) To c(2)(1), c(3)(0) To c(3)(1), _
                c(4)(0) To c(4)(1), c(5)(0) To c(5)(1), c(6)(0) To c(6)(1), c(7)(0) To c(7)(1), _
                c(8)(0) To c(8)(1), c(9)(0) To c(9)(1), c(10)(0) To c(10)(1), c(11)(0) To c(11)(1), _
                c(12)(0) To c(12)(1), c(13)(0) To c(13)(1))
        Case 14
            ReDim rtn(c(0)(0) To c(0)(1), c(1)(0) To c(1)(1), c(2)(0) To c(2)(1), c(3)(0) To c(3)(1), _
                c(4)(0) To c(4)(1), c(5)(0) To c(5)(1), c(6)(0) To c(6)(1), c(7)(0) To c(7)(1), _
                c(8)(0) To c(8)(1), c(9)(0) To c(9)(1), c(10)(0) To c(10)(1), c(11)(0) To c(11)(1), _
                c(12)(0) To c(12)(1), c(13)(0) To c(13)(1), c(14)(0) To c(14)(1))
        Case 15
            ReDim rtn(c(0)(0) To c(0)(1), c(1)(0) To c(1)(1), c(2)(0) To c(2)(1), c(3)(0) To c(3)(1), _
                c(4)(0) To c(4)(1), c(5)(0) To c(5)(1), c(6)(0) To c(6)(1), c(7)(0) To c(7)(1), _
                c(8)(0) To c(8)(1), c(9)(0) To c(9)(1), c(10)(0) To c(10)(1), c(11)(0) To c(11)(1), _
                c(12)(0) To c(12)(1), c(13)(0) To c(13)(1), c(14)(0) To c(14)(1), c(15)(0) To c(15)(1))
        Case 16
            ReDim rtn(c(0)(0) To c(0)(1), c(1)(0) To c(1)(1), c(2)(0) To c(2)(1), c(3)(0) To c(3)(1), _
                c(4)(0) To c(4)(1), c(5)(0) To c(5)(1), c(6)(0) To c(6)(1), c(7)(0) To c(7)(1), _
                c(8)(0) To c(8)(1), c(9)(0) To c(9)(1), c(10)(0) To c(10)(1), c(11)(0) To c(11)(1), _
                c(12)(0) To c(12)(1), c(13)(0) To c(13)(1), c(14)(0) To c(14)(1), c(15)(0) To c(15)(1), _
                c(16)(0) To c(16)(1))
        Case 17
            ReDim rtn(c(0)(0) To c(0)(1), c(1)(0) To c(1)(1), c(2)(0) To c(2)(1), c(3)(0) To c(3)(1), _
                c(4)(0) To c(4)(1), c(5)(0) To c(5)(1), c(6)(0) To c(6)(1), c(7)(0) To c(7)(1), _
                c(8)(0) To c(8)(1), c(9)(0) To c(9)(1), c(10)(0) To c(10)(1), c(11)(0) To c(11)(1), _
                c(12)(0) To c(12)(1), c(13)(0) To c(13)(1), c(14)(0) To c(14)(1), c(15)(0) To c(15)(1), _
                c(16)(0) To c(16)(1), c(17)(0) To c(17)(1))
        Case 18
            ReDim rtn(c(0)(0) To c(0)(1), c(1)(0) To c(1)(1), c(2)(0) To c(2)(1), c(3)(0) To c(3)(1), _
                c(4)(0) To c(4)(1), c(5)(0) To c(5)(1), c(6)(0) To c(6)(1), c(7)(0) To c(7)(1), _
                c(8)(0) To c(8)(1), c(9)(0) To c(9)(1), c(10)(0) To c(10)(1), c(11)(0) To c(11)(1), _
                c(12)(0) To c(12)(1), c(13)(0) To c(13)(1), c(14)(0) To c(14)(1), c(15)(0) To c(15)(1), _
                c(16)(0) To c(16)(1), c(17)(0) To c(17)(1), c(18)(0) To c(18)(1))
        Case 19
            ReDim rtn(c(0)(0) To c(0)(1), c(1)(0) To c(1)(1), c(2)(0) To c(2)(1), c(3)(0) To c(3)(1), _
                c(4)(0) To c(4)(1), c(5)(0) To c(5)(1), c(6)(0) To c(6)(1), c(7)(0) To c(7)(1), _
                c(8)(0) To c(8)(1), c(9)(0) To c(9)(1), c(10)(0) To c(10)(1), c(11)(0) To c(11)(1), _
                c(12)(0) To c(12)(1), c(13)(0) To c(13)(1), c(14)(0) To c(14)(1), c(15)(0) To c(15)(1), _
                c(16)(0) To c(16)(1), c(17)(0) To c(17)(1), c(18)(0) To c(18)(1), c(19)(0) To c(19)(1))
        Case 20
            ReDim rtn(c(0)(0) To c(0)(1), c(1)(0) To c(1)(1), c(2)(0) To c(2)(1), c(3)(0) To c(3)(1), _
                c(4)(0) To c(4)(1), c(5)(0) To c(5)(1), c(6)(0) To c(6)(1), c(7)(0) To c(7)(1), _
                c(8)(0) To c(8)(1), c(9)(0) To c(9)(1), c(10)(0) To c(10)(1), c(11)(0) To c(11)(1), _
                c(12)(0) To c(12)(1), c(13)(0) To c(13)(1), c(14)(0) To c(14)(1), c(15)(0) To c(15)(1), _
                c(16)(0) To c(16)(1), c(17)(0) To c(17)(1), c(18)(0) To c(18)(1), c(19)(0) To c(19)(1), _
                c(20)(0) To c(20)(1))
    End Select
    createMultidimensionalArray = rtn
End Function









