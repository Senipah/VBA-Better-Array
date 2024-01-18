Attribute VB_Name = "TestUtils"
Option Explicit
'@Folder("VBABetterArray.Tests.Utils")
'@IgnoreModule IIfSideEffect, AssignmentNotUsed, ProcedureNotUsed
'@IgnoreModule FunctionNotUsed

#If Mac Then
    Private Const Sep As String = "/"
#Else
    Private Const Sep As String = "\"
#End If

Public Function CSVDataOutputPath(Optional ByVal fileName As String = "output.csv") As String
    Const DATA_DIR As String = "csv_data"
    CSVDataOutputPath = JoinPath(ThisWorkbook.Path, DATA_DIR, fileName)
End Function

Public Sub PrintExpectedActualStringsToConsole(ByVal Expected As String, ByVal Actual As String)
    Debug.Print ConsoleHeader("Expected")
    Debug.Print Expected
    Debug.Print ConsoleHeader("Actual")
    Debug.Print Actual
End Sub

Public Function WrapQuoteUtil(Optional ByVal Source As String = vbNullString) As String
    Dim quoteChar As String
    quoteChar = chr(34)
    WrapQuoteUtil = quoteChar & Source & quoteChar
End Function

Public Sub ReadCSV(ByVal Arr As BetterArray, ByVal Path As String, ByVal fileName As String)
    Dim startTime As Single
    Dim endTime As Single
    Dim Filepath As String
    Filepath = JoinPath(Path, fileName)
    Debug.Print ConsoleHeader("Reading: " & fileName)
    startTime = Timer
    '@Ignore FunctionReturnValueDiscarded
    Arr.FromCSVFile Filepath
    endTime = Timer
    Debug.Print "Time taken: " & endTime - startTime
End Sub

Public Function JoinPath(ParamArray Args() As Variant) As String
    Dim i As Long
    Dim argv() As Variant
    argv = CVar(Args)
    For i = LBound(argv) To UBound(argv)
        argv(i) = TrimSeparator(argv(i))
    Next
    JoinPath = Strings.Join(argv, Sep)
End Function

Private Function TrimSeparator(ByVal Path As String) As String
    If Right$(Path, 1) = Sep Then
        TrimSeparator = Left$(Path, Len(Path) - 1)
    Else
        TrimSeparator = Path
    End If
End Function

'@Ignore ProcedureNotUsed
Private Function LastRow(ByVal Target As Worksheet, Optional ByVal columnNum As Long = 1) As Long
    LastRow = Target.Cells.Item(Target.Rows.Count, columnNum).End(xlUp).Row
End Function

'@Ignore ProcedureNotUsed
Private Function LastCol(ByVal Target As Worksheet, Optional ByVal rowNum As Long = 1) As Long
    LastCol = Target.Cells.Item(rowNum, Target.Columns.Count).End(xlToLeft).Column
End Function

Public Function ConsoleHeader(ByVal descriptor As String) As String
    Dim Result As String
    Dim corner As String
    Dim vertice As String
    Dim horizon As String
    corner = "+"
    vertice = "|"
    horizon = "-"
    Result = corner & String(Len(descriptor) + 2, horizon) & corner & vbCrLf _
           & vertice & " " & descriptor & " " & vertice & vbCrLf _
           & corner & String(Len(descriptor) + 2, horizon) & corner
    ConsoleHeader = Result
End Function

Public Function SectionHeader(ByVal descriptor As String) As String
    Dim Result As String
    Dim edge As String
    edge = "'"
    Result = edge & String(Len(descriptor) + 2, edge) & edge & vbCrLf _
           & edge & " " & descriptor & " " & edge & vbCrLf _
           & edge & String(Len(descriptor) + 2, edge) & edge
    SectionHeader = Result
End Function


Public Sub RatePerformance(ByVal manualTime As Double, ByVal betterArrayTime As Double)
    Const descriptor As String = "Time taken with "
    Const resultStart As String = "BetterArray is "
    Const resultEnd As String = " Than the manual method."
    Dim Diff As Double
    
    
    Diff = manualTime - betterArrayTime
    If Diff <> 0 And betterArrayTime <> 0 Then Diff = Diff / betterArrayTime
    Debug.Print descriptor & "manual method: " & manualTime
    Debug.Print descriptor & "BetterArray: " & betterArrayTime
    If Diff <> 0 Then
        Debug.Print resultStart _
                    & Format$(Abs(Diff), "Percent") _
                    & IIf(Diff > 0, " faster", " slower") _
                    & resultEnd
    Else
        Debug.Print "Effectively same speed."
    End If
End Sub

''''''''''''''''''''
' Helper Functions '
''''''''''''''''''''

Public Function SequenceEqualsMultiVsJagged( _
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
    SequenceEqualsMultiVsJagged = True
    Exit Function
ErrHandler:
    On Error GoTo 0
End Function

Public Function SequenceEquals_JaggedArray( _
        ByRef Expected() As Variant, _
        ByRef Actual() As Variant _
    ) As Boolean
    Dim i As Long
    On Error GoTo ErrHandler
    For i = LBound(Expected) To UBound(Expected)
        If IsArray(Expected(i)) Then
            Dim expectedChild() As Variant
            Dim actualChild() As Variant
            expectedChild = Expected(i)
            actualChild = Actual(i)
            If Not SequenceEquals_JaggedArray(expectedChild, actualChild) Then
                GoTo ErrHandler
            End If
        Else
            If Not ElementsAreEqual(Expected(i), Actual(i)) Then
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
        ByRef Expected() As Variant, _
        ByRef Actual As Object, _
        Optional ByVal Transposed As Boolean _
    ) As Boolean
    Dim i As Long
    Dim j As Long
    
    On Error GoTo ErrHandler
    
    If TypeName(Actual) <> "Range" Or Actual Is Nothing Then
        GoTo ErrHandler
    End If
    
    For i = 1 To Actual.Rows.Count
        For j = 1 To Actual.Columns.Count
            If Not ElementsAreEqual( _
                Expected(IIf(Transposed, j - 1, i - 1), IIf(Transposed, i - 1, j - 1)), _
                Actual.Cells.Item(i, j).Value _
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
        ByVal Expected As Variant, _
        ByVal Actual As Variant _
    ) As Boolean
Attribute ElementsAreEqual.VB_Description = "Compares two values for equality. Doesn't support multidimensional arrays."
    ' Using 13dp of precision for EPSILON rather than IEEE 754 standard of 2^-52
    ' some roundings in type conversions cause greater diffs than machine epsilon
    Const Epsilon As Double = 0.0000000000001
    Dim Result As Boolean
    Dim i As Long
    
    On Error GoTo ErrHandler
    If IsArray(Expected) Or IsArray(Actual) Then
        If IsArray(Expected) And IsArray(Actual) Then
            If LBound(Expected) = LBound(Actual) And _
                    UBound(Expected) = UBound(Actual) Then
                Dim CurrentlyEqual As Boolean
                CurrentlyEqual = True
                For i = LBound(Expected) To UBound(Actual)
                    If Not ElementsAreEqual(Expected(i), Actual(i)) Then
                        CurrentlyEqual = False
                        Exit For
                    End If
                Next
                Result = CurrentlyEqual
            End If
        End If
    ElseIf IsEmpty(Expected) Or IsEmpty(Actual) Then
        If IsEmpty(Expected) And IsEmpty(Actual) Then Result = True
    ElseIf IsObject(Expected) Or IsObject(Actual) Then
        If IsObject(Expected) And IsObject(Actual) Then
            If Expected Is Actual Then Result = True
        End If
    ElseIf IsNumeric(Expected) Or IsNumeric(Actual) Then
        If IsNumeric(Expected) And IsNumeric(Actual) Then
            Dim Diff As Double
            Diff = Abs(Expected - Actual)
            If Diff <= (IIf( _
                    Abs(Expected) < Abs(Actual), _
                    Abs(Actual), _
                    Abs(Expected) _
                ) * Epsilon) Then
                Result = True
            End If
        End If
    ElseIf Expected = Actual Then
        Result = True
    End If
    ElementsAreEqual = Result
    Exit Function
ErrHandler:
    ElementsAreEqual = False
End Function


'@Description("For Unit Tests only. No MD array support)"
Public Function arraysAreReversed( _
        ByRef Original() As Variant, _
        ByRef reversed() As Variant, _
        Optional ByVal Recurse As Boolean _
    ) As Boolean
    Dim i As Long
    Dim LocalUpperBound As Long
    Dim LocalLowerBound As Long
    Dim Result As Boolean
    
    On Error GoTo ErrHandler
    
    LocalUpperBound = UBound(Original)
    LocalLowerBound = LBound(Original)
    Result = True
    
    For i = LocalLowerBound To LocalUpperBound
        If IsArray(Original(i)) Then
            If IsArray(reversed(LocalUpperBound + LocalLowerBound - i)) Then
                Dim originalArray() As Variant
                Dim reversedArray() As Variant
                originalArray = Original(i)
                reversedArray = reversed(LocalUpperBound + LocalLowerBound - i)
                If Recurse Then
                    If Not arraysAreReversed(originalArray, reversedArray) Then
                        Result = False
                        Exit For
                    End If
                Else
                    If Not SequenceEquals_JaggedArray(originalArray, reversedArray) Then
                        arraysAreReversed = False
                        Exit For
                    End If
                End If
            Else
                Result = False
                Exit For
            End If
        Else
            If Not ElementsAreEqual( _
                    Original(i), _
                    reversed(LocalUpperBound + LocalLowerBound - i) _
                ) Then
                Result = False
                Exit For
            End If
        End If
    Next
    arraysAreReversed = Result
    Exit Function
ErrHandler:
    arraysAreReversed = False
End Function
