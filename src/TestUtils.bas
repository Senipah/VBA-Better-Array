Attribute VB_Name = "TestUtils"
Option Explicit
'@Folder("BetterArray.Tests.Utils")

#If Mac Then
    Private Const SEP As String = "/"
#Else
    Private Const SEP As String = "\"
#End If

Public Sub ReadCSV(ByVal Arr As BetterArray, ByVal Path As String, ByVal fileName As String)
    Dim startTime As Single
    Dim endTime As Single
    Dim filePath As String
    filePath = JoinPath(Path, fileName)
    Debug.Print formatDescriptor("Reading: " & fileName)
    startTime = Timer
    '@Ignore FunctionReturnValueDiscarded
    Arr.FromCSVFile filePath
    endTime = Timer
    Debug.Print "Time taken: " & endTime - startTime
End Sub

Public Function JoinPath(ParamArray args() As Variant) As String
    Dim i As Long
    Dim argv() As Variant
    argv = CVar(args)
    For i = LBound(argv) To UBound(argv)
        argv(i) = TrimSeparator(argv(i))
    Next
    JoinPath = Strings.Join(argv, SEP)
End Function

Private Function TrimSeparator(ByVal Path As String) As String
    If Right$(Path, 1) = SEP Then
        TrimSeparator = Left$(Path, Len(Path) - 1)
    Else
        TrimSeparator = Path
    End If
End Function

'@Ignore ProcedureNotUsed
Private Function lastRow(ByVal target As Worksheet, Optional ByVal columnNum As Long = 1) As Long
    lastRow = target.Cells.Item(target.Rows.count, columnNum).End(xlUp).Row
End Function

'@Ignore ProcedureNotUsed
Private Function lastCol(ByVal target As Worksheet, Optional ByVal rowNum As Long = 1) As Long
    lastCol = target.Cells.Item(rowNum, target.Columns.count).End(xlToLeft).column
End Function

Public Function formatDescriptor(ByVal descriptor As String) As String
    Dim result As String
    Dim corner As String
    Dim vertice As String
    Dim horizon As String
    corner = "+"
    vertice = "|"
    horizon = "-"
    result = corner & String(Len(descriptor) + 2, horizon) & corner & vbCrLf _
           & vertice & " " & descriptor & " " & vertice & vbCrLf _
           & corner & String(Len(descriptor) + 2, horizon) & corner
    formatDescriptor = result
End Function

Public Sub RatePerformance(ByVal manualTime As Double, ByVal betterArrayTime As Double)
    Const descriptor As String = "Time taken with "
    Const resultStart As String = "BetterArray is "
    Const resultEnd As String = " Than the manual method."
    Dim diff As Double
    
    
    diff = manualTime - betterArrayTime
    If diff <> 0 And betterArrayTime <> 0 Then diff = diff / betterArrayTime
    Debug.Print descriptor & "manual method: " & manualTime
    Debug.Print descriptor & "BetterArray: " & betterArrayTime
    If diff <> 0 Then
        Debug.Print resultStart _
                    & Format$(Abs(diff), "Percent") _
                    & IIf(diff > 0, " faster", " slower") _
                    & resultEnd
    Else
        Debug.Print "Effectively same speed."
    End If
End Sub

''''''''''''''''''''
' Helper Functions '
''''''''''''''''''''

Public Function SequenceEqualsMutiVsJagged( _
    ByRef multi() As Variant, _
    ByRef jagged() As Variant _
) As Boolean
    Dim i As Long
    Dim j As Long
    For i = LBound(multi, 1) To UBound(multi, 1)
        For j = LBound(multi, 2) To UBound(multi, 2)
            If multi(i, j) <> jagged(i)(j) Then
                GoTo ErrHandler
            End If
        Next
    Next
    
    On Error GoTo 0
    SequenceEqualsMutiVsJagged = True
    Exit Function
ErrHandler:
    On Error GoTo 0
End Function

Public Function SequenceEquals_JaggedArray( _
        ByRef expected() As Variant, _
        ByRef actual() As Variant _
    ) As Boolean
    Dim i As Long
    On Error GoTo ErrHandler
    For i = LBound(expected) To UBound(expected)
        If IsArray(expected(i)) Then
            Dim expectedChild() As Variant
            Dim actualChild() As Variant
            expectedChild = expected(i)
            actualChild = actual(i)
            If Not SequenceEquals_JaggedArray(expectedChild, actualChild) Then
                GoTo ErrHandler
            End If
        Else
            If Not ElementsAreEqual(expected(i), actual(i)) Then
                GoTo ErrHandler
            End If
        End If
    Next
    On Error GoTo 0
    SequenceEquals_JaggedArray = True
    Exit Function
ErrHandler:
    On Error GoTo 0
End Function


Public Function SequenceEquals_JaggedArrayVsRange( _
        ByRef expected() As Variant, _
        ByRef actual As Object, _
        Optional ByVal transposed As Boolean _
    ) As Boolean
    Dim i As Long
    Dim j As Long
    
    On Error GoTo ErrHandler
    
    If TypeName(actual) <> "Range" Or actual Is Nothing Then
        GoTo ErrHandler
    End If
    
    For i = 1 To actual.Rows.count
        For j = 1 To actual.Columns.count
            If Not ElementsAreEqual( _
                expected(IIf(transposed, j - 1, i - 1), IIf(transposed, i - 1, j - 1)), _
                actual.Cells.Item(i, j).value _
            ) Then
                GoTo ErrHandler
            End If
        Next
    Next
    On Error GoTo 0
    SequenceEquals_JaggedArrayVsRange = True
    Exit Function
ErrHandler:
    On Error GoTo 0
End Function

'@Description("Compares two values for equality. Doesn't support multidimensional arrays.")
Public Function ElementsAreEqual( _
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


'@Description("For Unit Tests only. No MD array support)"
Public Function arraysAreReversed( _
        ByRef original() As Variant, _
        ByRef reversed() As Variant, _
        Optional ByVal recurse As Boolean _
    ) As Boolean
    Dim i As Long
    Dim localUpperBound As Long
    Dim localLowerBound As Long
    Dim result As Boolean
    
    On Error GoTo ErrHandler
    
    localUpperBound = UBound(original)
    localLowerBound = LBound(original)
    result = True
    
    For i = localLowerBound To localUpperBound
        If IsArray(original(i)) Then
            If IsArray(reversed(localUpperBound + localLowerBound - i)) Then
                Dim originalArray() As Variant
                Dim reversedArray() As Variant
                originalArray = original(i)
                reversedArray = reversed(localUpperBound + localLowerBound - i)
                If recurse Then
                    If Not arraysAreReversed(originalArray, reversedArray) Then
                        result = False
                        Exit For
                    End If
                Else
                    If Not SequenceEquals_JaggedArray(originalArray, reversedArray) Then
                        arraysAreReversed = False
                        Exit For
                    End If
                End If
            Else
                result = False
                Exit For
            End If
        Else
            If Not ElementsAreEqual( _
                    original(i), _
                    reversed(localUpperBound + localLowerBound - i) _
                ) Then
                result = False
                Exit For
            End If
        End If
    Next
    arraysAreReversed = result
    Exit Function
ErrHandler:
    arraysAreReversed = False
End Function
