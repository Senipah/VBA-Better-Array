VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AssertClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const EPSILON As Double = 0.0000000000001

Public Sub Fail(Optional ByVal Message As String = "Assertion failed.")
    RecordTestFailure Message
End Sub

Public Sub IsTrue(ByVal Condition As Boolean, Optional ByVal Message As String = vbNullString)
    If Not Condition Then
        Fail IIf(Len(Message) > 0, Message, "Expected True but got False.")
    End If
End Sub

Public Sub IsFalse(ByVal Condition As Boolean, Optional ByVal Message As String = vbNullString)
    If Condition Then
        Fail IIf(Len(Message) > 0, Message, "Expected False but got True.")
    End If
End Sub

Public Sub IsNotNothing(ByRef Actual As Variant, Optional ByVal Message As String = vbNullString)
    If IsObject(Actual) Then
        If Actual Is Nothing Then
            Fail IIf(Len(Message) > 0, Message, "Expected object reference to be set.")
        End If
    ElseIf IsEmpty(Actual) Then
        Fail IIf(Len(Message) > 0, Message, "Expected value to be initialized.")
    End If
End Sub

Public Sub AreEqual(ByRef Expected As Variant, ByRef Actual As Variant, Optional ByVal Message As Variant)
    If Not ValuesAreEqual(Expected, Actual) Then
        Fail BuildMessage("Expected values to be equal." & _
            " Expected: " & VariantToString(Expected) & _
            ", Actual: " & VariantToString(Actual), Message)
    End If
End Sub

Public Sub AreSame(ByRef Expected As Variant, ByRef Actual As Variant, Optional ByVal Message As Variant)
    If IsObject(Expected) And IsObject(Actual) Then
        If Not (Expected Is Actual) Then
            Fail BuildMessage("Expected object references to be the same.", Message)
        End If
    Else
        AreEqual Expected, Actual, Message
    End If
End Sub

Public Sub AreNotSame(ByRef Expected As Variant, ByRef Actual As Variant, Optional ByVal Message As Variant)
    If IsObject(Expected) And IsObject(Actual) Then
        If Expected Is Actual Then
            Fail BuildMessage("Expected object references to be different.", Message)
        End If
    ElseIf ValuesAreEqual(Expected, Actual) Then
        Fail BuildMessage("Expected values to be different.", Message)
    End If
End Sub

Public Sub SequenceEquals(ByRef Expected As Variant, ByRef Actual As Variant, Optional ByVal Message As Variant)
    If Not ValuesAreEqual(Expected, Actual) Then
        Fail BuildMessage("Expected sequences to be equal.", Message)
    End If
End Sub

Public Sub SeqenceEquals(ByRef Expected As Variant, ByRef Actual As Variant, Optional ByVal Message As Variant)
    SequenceEquals Expected, Actual, Message
End Sub

Public Sub NotSequenceEquals(ByRef Expected As Variant, ByRef Actual As Variant, Optional ByVal Message As Variant)
    If ValuesAreEqual(Expected, Actual) Then
        Fail BuildMessage("Expected sequences to be different.", Message)
    End If
End Sub

Private Function ValuesAreEqual(ByRef Expected As Variant, ByRef Actual As Variant) As Boolean
    On Error GoTo SafeFail

    If IsArray(Expected) Or IsArray(Actual) Then
        If IsArray(Expected) And IsArray(Actual) Then
            ValuesAreEqual = ArraysAreEqual(Expected, Actual)
        Else
            ValuesAreEqual = False
        End If
    ElseIf IsObject(Expected) Or IsObject(Actual) Then
        If IsObject(Expected) And IsObject(Actual) Then
            ValuesAreEqual = (Expected Is Actual)
        Else
            ValuesAreEqual = False
        End If
    ElseIf IsNull(Expected) Or IsNull(Actual) Then
        ValuesAreEqual = (IsNull(Expected) And IsNull(Actual))
    ElseIf IsEmpty(Expected) Or IsEmpty(Actual) Then
        ValuesAreEqual = (IsEmpty(Expected) And IsEmpty(Actual))
    ElseIf IsNumeric(Expected) Or IsNumeric(Actual) Then
        If IsNumeric(Expected) And IsNumeric(Actual) Then
            Dim Diff As Double
            Dim RelativeScale As Double
            Diff = Abs(CDbl(Expected) - CDbl(Actual))
            RelativeScale = IIf(Abs(CDbl(Expected)) < Abs(CDbl(Actual)), Abs(CDbl(Actual)), Abs(CDbl(Expected)))
            If RelativeScale = 0 Then
                ValuesAreEqual = (Diff = 0)
            Else
                ValuesAreEqual = (Diff <= RelativeScale * EPSILON)
            End If
        Else
            ValuesAreEqual = False
        End If
    Else
        ValuesAreEqual = (Expected = Actual)
    End If
    Exit Function

SafeFail:
    ValuesAreEqual = False
End Function

Private Function ArraysAreEqual(ByRef Expected As Variant, ByRef Actual As Variant) As Boolean
    Dim ExpectedDims As Long
    Dim ActualDims As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long

    ExpectedDims = ArrayDimensions(Expected)
    ActualDims = ArrayDimensions(Actual)

    If ExpectedDims <> ActualDims Then Exit Function
    If ExpectedDims = 0 Then
        ArraysAreEqual = True
        Exit Function
    End If
    If Not ArrayBoundsMatch(Expected, Actual, ExpectedDims) Then Exit Function

    Select Case ExpectedDims
    Case 1
        For i = LBound(Expected, 1) To UBound(Expected, 1)
            If Not ValuesAreEqual(Expected(i), Actual(i)) Then Exit Function
        Next
    Case 2
        For i = LBound(Expected, 1) To UBound(Expected, 1)
            For j = LBound(Expected, 2) To UBound(Expected, 2)
                If Not ValuesAreEqual(Expected(i, j), Actual(i, j)) Then Exit Function
            Next
        Next
    Case 3
        For i = LBound(Expected, 1) To UBound(Expected, 1)
            For j = LBound(Expected, 2) To UBound(Expected, 2)
                For k = LBound(Expected, 3) To UBound(Expected, 3)
                    If Not ValuesAreEqual(Expected(i, j, k), Actual(i, j, k)) Then Exit Function
                Next
            Next
        Next
    Case Else
        Exit Function
    End Select

    ArraysAreEqual = True
End Function

Private Function ArrayBoundsMatch(ByRef Expected As Variant, ByRef Actual As Variant, ByVal Dimensions As Long) As Boolean
    Dim i As Long
    On Error GoTo SafeFail

    For i = 1 To Dimensions
        If LBound(Expected, i) <> LBound(Actual, i) Then Exit Function
        If UBound(Expected, i) <> UBound(Actual, i) Then Exit Function
    Next

    ArrayBoundsMatch = True
    Exit Function

SafeFail:
    ArrayBoundsMatch = False
End Function

Private Function ArrayDimensions(ByRef Source As Variant) As Long
    Dim i As Long
    On Error GoTo Done
    For i = 1 To 60
        Dim BoundValue As Long
        BoundValue = LBound(Source, i)
    Next
Done:
    ArrayDimensions = i - 1
    On Error GoTo 0
End Function

Private Function BuildMessage(ByVal DefaultMessage As String, Optional ByVal Message As Variant) As String
    Dim CustomMessage As String
    CustomMessage = GetCustomMessage(Message)
    If Len(CustomMessage) > 0 Then
        BuildMessage = CustomMessage
    Else
        BuildMessage = DefaultMessage
    End If
End Function

Private Function GetCustomMessage(Optional ByVal Message As Variant) As String
    If IsMissing(Message) Then Exit Function
    If VarType(Message) = vbString Then
        GetCustomMessage = CStr(Message)
    End If
End Function

Private Function VariantToString(ByRef Value As Variant) As String
    On Error GoTo SafeFail

    If IsArray(Value) Then
        VariantToString = ArrayDescriptor(Value)
    ElseIf IsObject(Value) Then
        If Value Is Nothing Then
            VariantToString = "Nothing"
        Else
            VariantToString = "<Object:" & TypeName(Value) & ">"
        End If
    ElseIf IsNull(Value) Then
        VariantToString = "Null"
    ElseIf IsEmpty(Value) Then
        VariantToString = "Empty"
    Else
        VariantToString = CStr(Value)
    End If
    Exit Function

SafeFail:
    VariantToString = "<unprintable>"
End Function

Private Function ArrayDescriptor(ByRef Source As Variant) As String
    Dim Dims As Long
    Dim i As Long
    Dim Parts() As String

    Dims = ArrayDimensions(Source)
    If Dims = 0 Then
        ArrayDescriptor = "<Array:Unallocated>"
        Exit Function
    End If

    ReDim Parts(1 To Dims)
    For i = 1 To Dims
        Parts(i) = "[" & LBound(Source, i) & "..." & UBound(Source, i) & "]"
    Next
    ArrayDescriptor = "<Array:" & Join(Parts, ",") & ">"
End Function
