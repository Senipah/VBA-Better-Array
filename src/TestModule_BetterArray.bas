Attribute VB_Name = "TestModule_BetterArray"
Attribute VB_Description = "Unit Tests for 'BetterArray.cls'"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")
'@ModuleDescription("Unit Tests for 'BetterArray.cls'")

'@IgnoreModule ProcedureNotUsed
'@IgnoreModule LineLabelNotUsed
'@IgnoreModule EmptyMethod

'Private Assert As Object
'Move to early bind
Private Assert As AssertClass

'Private Fakes As Object
'Move to early bind
'@Ignore VariableNotUsed
Private Fakes As FakesProvider

' Module level declaration of system under test
Private SUT As BetterArray
' Module level declaration of ArrayGenerator as used by most tests
Private Gen As ArrayGenerator

Private Const TEST_ARRAY_LENGTH As Long = 10
Private Const EXCEL_DEPENDENCY_WARNING As String = "A test depending on an ExcelProvider instance had failed." & _
        " Once resolved ensure to end any orphan Excel processes running on this system."

' TODO: Expected error tests for all public methods with custom errors

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
'    Set Assert = CreateObject("Rubberduck.AssertClass")
'    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    ' Move to early binding
    Set Assert = New AssertClass
    Set Fakes = New FakesProvider
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'this method runs before every test in the module.
    Set SUT = New BetterArray
    Set Gen = New ArrayGenerator
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    Set SUT = Nothing
    Set Gen = Nothing
End Sub

''''''''''''''''''''
' Helper Functions '
''''''''''''''''''''

Private Function SequenceEquals_JaggedArray( _
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
            If Not valuesAreEqual(expected(i), actual(i)) Then
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


Private Function SequenceEquals_JaggedArrayVsRange( _
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
            If Not valuesAreEqual( _
                expected(IIf(transposed, j - 1, i - 1), IIf(transposed, i - 1, j - 1)), _
                actual.Cells.Item(i, j).Value _
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


Private Function valuesAreEqual(ByVal expected As Variant, ByVal actual As Variant) As Boolean
    ' Using 13dp of precision for EPSILON rather than IEEE 754 standard of 2^-52
    ' some roundings in type conversions cause greater diffs than machine epsilon
    Const Epsilon As Double = 0.0000000000001
    Dim result As Boolean
    Dim diff As Double
    If IsEmpty(expected) Then
        If IsEmpty(actual) Then result = True
    ElseIf IsNumeric(expected) Then
        diff = Abs(expected - actual)
        If diff <= (IIf(Abs(expected) < Abs(actual), Abs(actual), Abs(expected)) * Epsilon) Then
            result = True
        End If
    ElseIf expected = actual Then
        result = True
    End If
    valuesAreEqual = result
End Function


'''''''''''''''''
' Instantiation '
'''''''''''''''''

'@TestMethod("BetterArray_Constructor")
Private Sub Constructor_CanInstantiate_SUTNotNothing()
    On Error GoTo TestFail
    
    'Arrange:
    'Act:
    'Assert:
    Assert.IsNotNothing SUT

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_Constructor")
Private Sub Constructor_CreatesWithDefaultCapacity_CapacityIsFour()
    On Error GoTo TestFail

    'Arrange:
    Const expected As Long = 4
    Dim actual As Long

    'Act:
    actual = SUT.Capacity

    'Assert:
    Assert.AreEqual expected, actual, "Default capacity incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


''''''''''''''''''''''''''''''''''''
' Attribute - DefaultMember - Item '
''''''''''''''''''''''''''''''''''''

'@TestMethod("BetterArray_Items")
Private Sub Items_DefaultMember_DefaultMemberAccessReturnsItems()
    On Error GoTo TestFail

    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim i As Long

    expected = Gen.GetArray()
    
    'Act:
    For i = LBound(expected) To UBound(expected)
        '@Ignore IndexedDefaultMemberAccess
        SUT(i) = expected(i)
    Next
    actual = SUT.Items

    'Assert:
    Assert.SequenceEquals expected, actual, "Actual array does not match expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'''''''''''''''''''
' Prop - Capacity '
'''''''''''''''''''

'@TestMethod("BetterArray_Capacity")
'@Ignore DuplicatedAnnotation
Private Sub Capacity_CanSetCapacity_ReturnedCapacityMatchesSetCapacity()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As Long = 20
    Dim actual As Long
   
    'Act:
    SUT.Capacity = expected
    actual = SUT.Capacity

    'Assert:
    Assert.AreEqual expected, actual, "Returned capacity does not equal set capacity"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


''''''''''''''''
' Prop - Items '
''''''''''''''''

'@TestMethod("BetterArray_Items")
Private Sub Items_CanAssignOneDimemsionalArray_ReturnedArrayEqualsAssignedArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant

    expected = Gen.GetArray(AG_VARIANT, AG_ONEDIMENSION)
   
    'Act:
    SUT.Items = expected
    actual = SUT.Items

    'Assert:
    Assert.SequenceEquals expected, actual, "Actual array does not match expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Items")
Private Sub Items_CanAssignMultiDimemsionalArray_ReturnedArrayEqualsAssignedArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    
    expected = Gen.GetArray(AG_VARIANT, AG_MULTIDIMENSION)
 
    'Act:
    SUT.Items = expected
    actual = SUT.Items

    'Assert:
    Assert.SequenceEquals expected, actual, "Actual array does not match expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Items")
' NOTE: does not use SequenceEquals due to Rubberduck issue: https://github.com/rubberduck-vba/Rubberduck/issues/5161
Private Sub Items_CanAssignJaggedArray_ReturnedArrayEqualsAssignedArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testResult As Boolean

    expected = Gen.GetArray(AG_VARIANT, AG_JAGGED)
    
    'Act:
    SUT.Items = expected
    actual = SUT.Items
    
    testResult = SequenceEquals_JaggedArray(expected, actual)

    'Assert:
    Assert.IsTrue testResult, "Contents of expected and actual do not match"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'''''''''''''''''
' Prop - Length '
'''''''''''''''''

'@TestMethod("BetterArray_Length")
Private Sub Length_FromAssignedOneDimensionalArray_ReturnedLengthEqualsOriginalArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim expected As Long
    Dim actual As Long
    
    expected = TEST_ARRAY_LENGTH
    testArray = Gen.GetArray(AG_VARIANT, AG_ONEDIMENSION)
    
    'Act:
    SUT.Items = testArray
    actual = SUT.Length

    'Assert:
    Assert.AreEqual expected, actual, "Actual length <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Length")
Private Sub Length_FromAssignedMultiDimensionalArray_ReturnedLengthEqualsOriginalArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim expected As Long
    Dim actual As Long

    expected = TEST_ARRAY_LENGTH
    testArray = Gen.GetArray(AG_VARIANT, AG_MULTIDIMENSION)

    'Act:
    SUT.Items = testArray
    actual = SUT.Length
    
    'Assert:
    Assert.AreEqual expected, actual, "Actual length <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Length")
Private Sub Length_FromAssignedJaggedArray_ReturnedLengthEqualsOriginalArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim expected As Long
    Dim actual As Long

    expected = TEST_ARRAY_LENGTH
    testArray = Gen.GetArray(AG_VARIANT, AG_JAGGED)

    'Act:
    SUT.Items = testArray
    actual = SUT.Length

    'Assert:
    Assert.AreEqual expected, actual, "Actual length <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_Length")
Private Sub Upperbound_FromAssignedOneDimensionalArray_ReturnedUpperBoundEqualsOriginalArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim expected As Long
    Dim actual As Long

    testArray = Gen.GetArray(AG_VARIANT, AG_ONEDIMENSION)
    expected = UBound(testArray)

    'Act:
    SUT.Items = testArray
    actual = SUT.UpperBound

    'Assert:
    Assert.AreEqual expected, actual, "Actual upperbound <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'''''''''''''''
' Prop - LowerBound '
'''''''''''''''


'@TestMethod("BetterArray_LowerBound")
Private Sub LowerBound_FromAssignedOneDimensionalArray_ReturnedLowerBoundEqualsOriginalArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim expected As Long
    Dim actual As Long

    testArray = Gen.GetArray(AG_VARIANT, AG_ONEDIMENSION)
    expected = LBound(testArray)
    
    'Act:
    SUT.Items = testArray
    actual = SUT.LowerBound

    'Assert:
    Assert.AreEqual expected, actual, "Actual LowerBound <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_LowerBound")
Private Sub LowerBound_ChangingLowerBoundOfAssignedArray_ReturnedArrayHasNewLowerBound()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim returnedItems As Variant
    Dim expected As Long
    Dim actual As Long
    Dim oldLowerBound As Long
    
    testArray = Gen.GetArray()
    oldLowerBound = LBound(testArray)
        
    'Act:
    SUT.Items = testArray
    expected = oldLowerBound + 1
    SUT.LowerBound = expected
    returnedItems = SUT.Items
    actual = LBound(returnedItems)

    'Assert:
    Assert.AreEqual expected, actual, "Actual LowerBound <> expected"
    Assert.AreEqual SUT.LowerBound, actual, "Actual LowerBound <> SUT.LowerBound prop"
    Assert.AreEqual UBound(testArray) + 1, UBound(returnedItems), "Actual upperbound <> expected"
    Assert.AreEqual SUT.UpperBound, UBound(returnedItems), "Actual upperbound <> SUT.UpperBound prop"
    Assert.AreEqual SUT.Length, TEST_ARRAY_LENGTH, "Actual length does not equal expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'''''''''''''''
' Prop - Item '
'''''''''''''''

'@TestMethod("BetterArray_Item")
Private Sub Item_ChangingExistingIndex_ItemIsChanged()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As String = "Hello World"
    Dim testArray() As Variant
    Dim actual As Variant
    Dim actualLowerBound As Long
    Dim expectedLowerBound As Long

    testArray = Gen.GetArray(AG_VARIANT, AG_ONEDIMENSION)
    expectedLowerBound = LBound(testArray)
    
    'Act:
    SUT.Items = testArray
    SUT.Item(1) = expected
    actual = SUT.Item(1)
    actualLowerBound = SUT.LowerBound

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
    Assert.AreEqual expectedLowerBound, actualLowerBound, "Actual LowerBound does not equal expected LowerBound"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Item")
Private Sub Item_ChangingIndexOverUpperBound_ItemIsPushed()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As String = "Hello World"
    Dim testArray() As Variant
    Dim actual As Variant
    Dim actualLowerBound As Long
    Dim expectedLowerBound As Long
    
    testArray = Gen.GetArray(AG_VARIANT, AG_ONEDIMENSION)
    expectedLowerBound = LBound(testArray)
    
    'Act:
    SUT.Items = testArray
    SUT.Item(SUT.UpperBound + 1) = expected
    actual = SUT.Item(SUT.UpperBound)
    actualLowerBound = SUT.LowerBound
    
    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
    Assert.AreEqual TEST_ARRAY_LENGTH + 1, SUT.Length, "Actual length does not match expected length"
    Assert.AreEqual expectedLowerBound, actualLowerBound, "Actual LowerBound does not match expected LowerBound"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Item")
Private Sub Item_ChangingIndexBelowLowerBound_ItemIsUnshifted()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As String = "Hello World"
    Dim testArray() As Variant
    Dim actual As Variant
    Dim expectedLowerBound As Long
    Dim actualLowerBound As Long

    testArray = Gen.GetArray(AG_VARIANT, AG_ONEDIMENSION)
    expectedLowerBound = LBound(testArray)
    
    'Act:
    SUT.Items = testArray
    SUT.Item(SUT.LowerBound - 10) = expected
    actual = SUT.Item(SUT.LowerBound)
    actualLowerBound = SUT.LowerBound

    'Assert:
    Assert.AreEqual expected, actual, "Actual result does not match expected result"
    Assert.AreEqual TEST_ARRAY_LENGTH + 1, SUT.Length, "Actual length does not match expected length"
    Assert.AreEqual expectedLowerBound, actualLowerBound, "Actual LowerBound does not match expected LowerBound"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_Item")
Private Sub Item_GetScalarValue_ValueReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim expected As Variant
    Dim actual As Variant
       
    testArray = Gen.GetArray(AG_VARIANT, AG_ONEDIMENSION)
    expected = testArray(1)
    
    'Act:
    SUT.Items = testArray
    actual = SUT.Item(1)

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_Item")
Private Sub Item_GetObject_SameObjectReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim expected As Object
    Dim actual As Object

    testArray = Gen.GetArray(AG_OBJECT, AG_ONEDIMENSION)
    Set expected = testArray(1)
    
    'Act:
    SUT.Items = testArray
    Set actual = SUT.Item(1)

    'Assert:
    Assert.AreSame expected, actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'''''''''''''''''
' Method - Push '
'''''''''''''''''

'@TestMethod("BetterArray_Push")
Private Sub Push_AddToNewBetterArray_ItemAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As String = "Hello World"
    Const expectedLength As Long = 1
    Const expectedUpperBound As Long = 0

    Dim actual As String
    Dim actualLength As Long
    Dim actualUpperBound As Long
    
    'Act:
    SUT.Push expected
    actual = SUT.Item(SUT.LowerBound)
    actualLength = SUT.Length
    actualUpperBound = SUT.UpperBound

    'Assert:
    Assert.AreEqual expected, actual, "actual <> expected"
    Assert.AreEqual expectedLength, actualLength, "Length value incorrect"
    Assert.AreEqual expectedUpperBound, actualUpperBound, "Upperbound value incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Push")
Private Sub Push_AddToExistingOneDimensionalArray_ItemAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As String = "Hello World"

    Dim testArray() As Variant
    Dim actual As String
    
    testArray = Gen.GetArray(AG_VARIANT, AG_ONEDIMENSION)
    
    'Act:
    SUT.Items = testArray
    SUT.Push expected
    actual = SUT.Item(SUT.UpperBound)

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
    Assert.AreEqual TEST_ARRAY_LENGTH + 1, SUT.Length, "Length value incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Push")
Private Sub Push_AddToExistingMultidimensionalArray_ItemAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Variant
    Dim actual As Variant
    Dim testArray() As Variant
    Dim returnedArray() As Variant

    expected = "Hello World"
    testArray = Gen.GetArray(AG_VARIANT, AG_MULTIDIMENSION)
    
    'Act:
    SUT.Items = testArray
    SUT.Push expected
    returnedArray = SUT.Items
    actual = returnedArray(UBound(returnedArray), LBound(returnedArray, 2))

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
    Assert.AreEqual TEST_ARRAY_LENGTH + 1, SUT.Length, "Length value incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_Push")
Private Sub Push_AddToExistingJaggedArray_ItemAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Variant
    Dim actual As Variant
    Dim testArray() As Variant
    Dim returnedArray() As Variant

    expected = Gen.GetArray(AG_VARIANT, AG_ONEDIMENSION)
    testArray = Gen.GetArray(AG_VARIANT, AG_JAGGED)
    
    'Act:
    SUT.Items = testArray
    SUT.Push expected
    returnedArray = SUT.Items
    actual = returnedArray(UBound(returnedArray))
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Element value incorrect"
    Assert.AreEqual TEST_ARRAY_LENGTH + 1, SUT.Length, "Length value incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_Push")
Private Sub Push_AddMultipleToNewBetterArray_ItemsAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As Long = 1
    Const expectedLength As Long = 3
    Const expectedUpperBound As Long = 2

    Dim actual As Long
    Dim actualLength As Long
    Dim actualUpperBound As Long
        
    'Act:
    SUT.Push expected, 2, 3
    actual = SUT.Item(SUT.LowerBound)
    actualLength = SUT.Length
    actualUpperBound = SUT.UpperBound

    'Assert:
    Assert.AreEqual expected, actual, "Element value incorrect"
    Assert.AreEqual expectedLength, actualLength, "Length value incorrect"
    Assert.AreEqual expectedUpperBound, actualUpperBound, "Upperbound value incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


''''''''''''''''
' Method - Pop '
''''''''''''''''

'@TestMethod("BetterArray_Pop")
Private Sub Pop_OneDimensionalArray_LastItemRemoved()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim actualLowerBound As Long
    Dim expectedLowerBound As Long
    Dim expected As String
    Dim actual As String
    
    testArray = Gen.GetArray(AG_STRING, AG_ONEDIMENSION)
    expected = testArray(UBound(testArray))
    expectedLowerBound = SUT.LowerBound
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Pop
    actualLowerBound = SUT.LowerBound

    'Assert:
    Assert.AreEqual expected, actual, "Element value incorrect"
    Assert.AreEqual TEST_ARRAY_LENGTH - 1, SUT.Length, "Length value incorrect"
    Assert.AreEqual UBound(testArray) - 1, SUT.UpperBound, "Upperbound value incorrect"
    Assert.AreEqual expectedLowerBound, actualLowerBound, "LowerBound value incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_Pop")
Private Sub Pop_ArrayLengthIsZero_ReturnsEmpty()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Variant
    Dim expectedLowerBound As Long
    Dim expectedLength As Long
    Dim expectedUpperBound As Long
    Dim actual As Variant
    Dim actualLowerBound As Long
    Dim actualLength As Long
    Dim actualUpperBound As Long
    
    expected = Empty
    expectedLowerBound = 0
    expectedLength = 0
    expectedUpperBound = -1
    
    'Act:
    actual = SUT.Pop
    actualLowerBound = SUT.LowerBound
    actualLength = SUT.Length
    actualUpperBound = SUT.UpperBound

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
    Assert.AreEqual expectedLowerBound, actualLowerBound, "LowerBound value incorrect"
    Assert.AreEqual expectedLength, actualLength, "Length value incorrect"
    Assert.AreEqual expectedUpperBound, actualUpperBound, "Upperbound value incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


''''''''''''''''''
' Method - Shift '
''''''''''''''''''

'@TestMethod("BetterArray_Shift")
Private Sub Shift_OneDimensionalArray_FirstItemRemoved()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim actualLowerBound As Long
    Dim expectedLowerBound As Long
    Dim expected As String
    Dim actual As String

    testArray = Gen.GetArray(AG_STRING, AG_ONEDIMENSION)
    expected = testArray(LBound(testArray))
    expectedLowerBound = SUT.LowerBound
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Shift
    actualLowerBound = SUT.LowerBound

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
    Assert.AreEqual TEST_ARRAY_LENGTH - 1, SUT.Length, "Length value incorrect"
    Assert.AreEqual UBound(testArray) - 1, SUT.UpperBound, "Upperbound value incorrect"
    Assert.AreEqual expectedLowerBound, actualLowerBound, "LowerBound value incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Shift")
Private Sub Shift_ArrayLengthIsZero_ReturnsEmpty()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As Variant = Empty
    Const expectedLowerBound As Long = 0
    Const expectedLength As Long = 0
    Const expectedUpperBound As Long = -1

    Dim actual As Variant
    Dim actualLowerBound As Long
    Dim actualLength As Long
    Dim actualUpperBound As Long
    
    'Act:
    actual = SUT.Shift
    actualLowerBound = SUT.LowerBound
    actualLength = SUT.Length
    actualUpperBound = SUT.UpperBound

    'Assert:
    Assert.AreEqual expected, actual, "Element value incorrect"
    Assert.AreEqual expectedLowerBound, actualLowerBound, "LowerBound value incorrect"
    Assert.AreEqual expectedLength, actualLength, "Length value incorrect"
    Assert.AreEqual expectedUpperBound, actualUpperBound, "Upperbound value incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

''''''''''''''''''''
' Method - Unshift '
''''''''''''''''''''

'@TestMethod("BetterArray_Unshift")
Private Sub Unshift_OneDimensionalArray_ItemAddedToBeginning()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim expected As String
    Dim actual As String
    Dim actualLowerBound As Long
    Dim expectedLowerBound As Long
    Dim testElement As String
    
    testArray = Gen.GetArray(AG_STRING, AG_ONEDIMENSION)
    testElement = "Hello World"
    expectedLowerBound = SUT.LowerBound
    expected = TEST_ARRAY_LENGTH + 1
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Unshift(testElement)
    actualLowerBound = SUT.LowerBound

    'Assert:
    Assert.AreEqual expected, actual, "Return value incorrect"
    Assert.AreEqual (UBound(testArray) + 1), SUT.UpperBound, "Upperbound value incorrect"
    Assert.AreEqual expectedLowerBound, actualLowerBound, "LowerBound value incorrect"
    Assert.AreEqual testElement, SUT.Item(SUT.LowerBound), "Element not inserted at correct position"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_Unshift")
Private Sub Unshift_ArrayLengthIsZero_ItemIsPushedToEmptyArray()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As Long = 1
    Const expectedLowerBound As Long = 0
    Const expectedUpperBound As Long = 0
    Const expectedElement As String = "Hello World"

    Dim actual As Long
    Dim actualLowerBound As Long
    Dim actualUpperBound As Long
    Dim actualElement As String
    
    'Act:
    actual = SUT.Unshift(expectedElement)
    actualLowerBound = SUT.LowerBound
    actualUpperBound = SUT.UpperBound
    actualElement = SUT.Item(SUT.LowerBound)

    'Assert:
    Assert.AreEqual expected, actual, "Actual length <> expected length"
    Assert.AreEqual expectedUpperBound, actualUpperBound, "Upperbound value incorrect"
    Assert.AreEqual expectedLowerBound, actualLowerBound, "LowerBound value incorrect"
    Assert.AreEqual expectedElement, actualElement, "Actual element <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_Unshift")
Private Sub Unshift_MultidimensionalArray_ItemAddedToBeginning()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As Long = TEST_ARRAY_LENGTH + 1
    Const expectedLowerBound As Long = 0
    Const expectedUpperBound As Long = TEST_ARRAY_LENGTH
    Const expectedElement As String = "Hello World"

    Dim actual As Long
    Dim actualLowerBound As Long
    Dim actualUpperBound As Long
    Dim actualElement As String
    Dim testArray() As Variant
    Dim returnedItems() As Variant

    testArray = Gen.GetArray(AG_VARIANT, AG_MULTIDIMENSION)
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Unshift(expectedElement)
    returnedItems = SUT.Items
    actualLowerBound = SUT.LowerBound
    actualUpperBound = SUT.UpperBound
    actualElement = returnedItems(LBound(returnedItems), LBound(returnedItems, 2))

    'Assert:
    Assert.AreEqual expected, actual, "Actual length <> expected length"
    Assert.AreEqual expectedUpperBound, actualUpperBound, "Upperbound value incorrect"
    Assert.AreEqual expectedLowerBound, actualLowerBound, "LowerBound value incorrect"
    Assert.AreEqual expectedElement, actualElement, "Actual element <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'''''''''''''''''''
' Method - Concat '
'''''''''''''''''''
'TODO: Concat test cases

'@TestMethod("BetterArray_Concat")
Private Sub Concat_AddOneDimArrayToEmptyInternal_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim expectedLength As Long
    Dim actualLength As Long
    Dim expectedUpperBound As Long
    Dim actualUpperBound As Long
    
    expectedLength = TEST_ARRAY_LENGTH
    expected = Gen.GetArray(Length:=expectedLength)
    expectedUpperBound = UBound(expected)
    
    'Act:
    actual = SUT.Concat(expected).Items
    actualLength = SUT.Length
    actualUpperBound = SUT.UpperBound
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> Expected"
    Assert.AreEqual expectedLength, actualLength, "Actual length <> expected"
    Assert.AreEqual expectedUpperBound, actualUpperBound, "Actual UpperBound <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Concat")
Private Sub Concat_AddMultipleOneDimArrayToEmptyInternal_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim expectedLength As Long
    Dim actualLength As Long
    Dim expectedUpperBound As Long
    Dim actualUpperBound As Long
    Dim firstAray() As Variant
    Dim secondArray() As Variant
    
    firstAray = Array(1, 2, 3)
    secondArray = Array(4, 5, 6)
    expected = Array(1, 2, 3, 4, 5, 6)
    expectedLength = 6
    expectedUpperBound = UBound(expected)
    
    'Act:
    actual = SUT.Concat(firstAray, secondArray).Items
    actualLength = SUT.Length
    actualUpperBound = SUT.UpperBound
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> Expected"
    Assert.AreEqual expectedLength, actualLength, "Actual length <> expected"
    Assert.AreEqual expectedUpperBound, actualUpperBound, "Actual UpperBound <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Concat")
Private Sub Concat_AddOneDimArrayToExistingOneDimArray_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim expectedLength As Long
    Dim actualLength As Long
    Dim expectedUpperBound As Long
    Dim actualUpperBound As Long
    Dim firstArray() As Variant
    Dim secondArray() As Variant
    
    firstArray = Gen.GetArray()
    secondArray = Gen.GetArray()
    expected = Gen.ConcatArraysOfSameStructure(AG_ONEDIMENSION, firstArray, secondArray)
    expectedLength = Gen.GetArrayLength(expected)
    expectedUpperBound = UBound(expected)
    
    'Act:
    SUT.Items = firstArray
    actual = SUT.Concat(secondArray).Items
    actualLength = SUT.Length
    actualUpperBound = SUT.UpperBound
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> Expected"
    Assert.AreEqual expectedLength, actualLength, "Actual length <> expected"
    Assert.AreEqual expectedUpperBound, actualUpperBound, "Actual UpperBound <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Concat")
Private Sub Concat_AddMultiDimArrayToEmptyInternal_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim expectedLength As Long
    Dim actualLength As Long
    Dim expectedUpperBound As Long
    Dim actualUpperBound As Long
    
    expectedLength = TEST_ARRAY_LENGTH
    expected = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION, Length:=expectedLength)
    expectedUpperBound = UBound(expected)
    
    'Act:
    actual = SUT.Concat(expected).Items
    actualLength = SUT.Length
    actualUpperBound = SUT.UpperBound
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> Expected"
    Assert.AreEqual expectedLength, actualLength, "Actual length <> expected"
    Assert.AreEqual expectedUpperBound, actualUpperBound, "Actual UpperBound <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Concat")
Private Sub Concat_AddMultiDimArrayToExistingMultiDimArray_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim firstArray() As Variant
    Dim secondArray() As Variant
    Dim actual() As Variant
    Dim expectedLength As Long
    Dim actualLength As Long
    Dim expectedUpperBound As Long
    Dim actualUpperBound As Long
    
    expectedLength = TEST_ARRAY_LENGTH * 2
    firstArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    secondArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    expected = Gen.ConcatArraysOfSameStructure(AG_MULTIDIMENSION, firstArray, secondArray)
    expectedUpperBound = UBound(expected)
    
    'Act:
    SUT.Items = firstArray
    actual = SUT.Concat(secondArray).Items
    actualLength = SUT.Length
    actualUpperBound = SUT.UpperBound
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> Expected"
    Assert.AreEqual expectedLength, actualLength, "Actual length <> expected"
    Assert.AreEqual expectedUpperBound, actualUpperBound, "Actual UpperBound <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Concat")
Private Sub Concat_AddJaggedArrayToEmptyInternal_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim expectedLength As Long
    Dim actualLength As Long
    Dim expectedUpperBound As Long
    Dim actualUpperBound As Long
    Dim testResult As Boolean
    
    expected = Gen.GetArray(ArrayType:=AG_JAGGED)
    expectedLength = Gen.GetArrayLength(expected)
    expectedUpperBound = UBound(expected)
    
    'Act:
    actual = SUT.Concat(expected).Items
    actualLength = SUT.Length
    actualUpperBound = SUT.UpperBound
    testResult = SequenceEquals_JaggedArray(expected, actual)
    
    'Assert:
    Assert.IsTrue testResult, "Actual <> Expected"
    Assert.AreEqual expectedLength, actualLength, "Actual length <> expected"
    Assert.AreEqual expectedUpperBound, actualUpperBound, "Actual UpperBound <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Concat")
Private Sub Concat_AddJaggedArrayToExistingJagged_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim firstArray() As Variant
    Dim secondArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    Dim expectedLength As Long
    Dim actualLength As Long
    Dim expectedUpperBound As Long
    Dim actualUpperBound As Long
    Dim testResult As Boolean
    
    firstArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    secondArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    
    expected = Gen.ConcatArraysOfSameStructure(AG_JAGGED, firstArray, secondArray)
    expectedLength = Gen.GetArrayLength(expected)
    expectedUpperBound = UBound(expected)
    
    'Act:
    SUT.Items = firstArray
    actual = SUT.Concat(secondArray).Items
    actualLength = SUT.Length
    actualUpperBound = SUT.UpperBound
    testResult = SequenceEquals_JaggedArray(expected, actual)
    
    'Assert:
    Assert.IsTrue testResult, "Actual <> Expected"
    Assert.AreEqual expectedLength, actualLength, "Actual length <> expected"
    Assert.AreEqual expectedUpperBound, actualUpperBound, "Actual UpperBound <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Concat")
Private Sub Concat_AddOneDimArrayToExistingJagged_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim firstArray() As Variant
    Dim secondArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    Dim expectedLength As Long
    Dim actualLength As Long
    Dim expectedUpperBound As Long
    Dim actualUpperBound As Long
    Dim testResult As Boolean
    
    firstArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    secondArray = Gen.GetArray(ArrayType:=AG_ONEDIMENSION)
    
    expected = Gen.ConcatArraysOfSameStructure(AG_JAGGED, firstArray, secondArray)
    expectedLength = Gen.GetArrayLength(expected)
    expectedUpperBound = UBound(expected)
    
    'Act:
    SUT.Items = firstArray
    actual = SUT.Concat(secondArray).Items
    actualLength = SUT.Length
    actualUpperBound = SUT.UpperBound
    testResult = SequenceEquals_JaggedArray(expected, actual)
    
    'Assert:
    Assert.IsTrue testResult, "Actual <> Expected"
    Assert.AreEqual expectedLength, actualLength, "Actual length <> expected"
    Assert.AreEqual expectedUpperBound, actualUpperBound, "Actual UpperBound <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Concat")
Private Sub Concat_AddOneDimArrayToExistingMulti_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim firstArray() As Variant
    Dim secondArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    Dim expectedLength As Long
    Dim actualLength As Long
    Dim expectedUpperBound As Long
    Dim actualUpperBound As Long
    
    ReDim firstArray(1 To 2, 1 To 2)
    firstArray(1, 1) = "Foo"
    firstArray(1, 2) = "Bar"
    firstArray(2, 1) = "Fizz"
    firstArray(2, 2) = "Buzz"
    
    secondArray = Array(1, 2, 3)
    
    ReDim expected(1 To 5, 1 To 2)
    expected(1, 1) = firstArray(1, 1)
    expected(1, 2) = firstArray(1, 2)
    expected(2, 1) = firstArray(2, 1)
    expected(2, 2) = firstArray(2, 2)
    expected(3, 1) = secondArray(0)
    expected(3, 2) = Empty
    expected(4, 1) = secondArray(1)
    expected(4, 2) = Empty
    expected(5, 1) = secondArray(2)
    expected(5, 2) = Empty
    
    expectedLength = 5
    expectedUpperBound = 5
    
    'Act:
    SUT.Items = firstArray
    actual = SUT.Concat(secondArray).Items
    actualLength = SUT.Length
    actualUpperBound = SUT.UpperBound
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> Expected"
    Assert.AreEqual expectedLength, actualLength, "Actual length <> expected"
    Assert.AreEqual expectedUpperBound, actualUpperBound, "Actual UpperBound <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
    
End Sub

'@TestMethod("BetterArray_Concat")
Private Sub Concat_AddMultiDimArrayToExistingOneDimArray_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim firstArray() As Variant
    Dim secondArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    Dim expectedLength As Long
    Dim actualLength As Long
    Dim expectedUpperBound As Long
    Dim actualUpperBound As Long
    
    firstArray = Array(1, 2, 3)
    ReDim secondArray(1 To 2, 1 To 2)
    secondArray(1, 1) = "Foo"
    secondArray(1, 2) = "Bar"
    secondArray(2, 1) = "Fizz"
    secondArray(2, 2) = "Buzz"
    
    ReDim expected(0 To 4, 0 To 1)
    expected(0, 0) = firstArray(0)
    expected(0, 1) = Empty
    expected(1, 0) = firstArray(1)
    expected(1, 1) = Empty
    expected(2, 0) = firstArray(2)
    expected(2, 1) = Empty
    expected(3, 0) = secondArray(1, 1)
    expected(3, 1) = secondArray(1, 2)
    expected(4, 0) = secondArray(2, 1)
    expected(4, 1) = secondArray(2, 2)
    
    expectedLength = 5
    expectedUpperBound = 4
    
    'Act:
    SUT.Items = firstArray
    actual = SUT.Concat(secondArray).Items
    actualLength = SUT.Length
    actualUpperBound = SUT.UpperBound
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> Expected"
    Assert.AreEqual expectedLength, actualLength, "Actual length <> expected"
    Assert.AreEqual expectedUpperBound, actualUpperBound, "Actual UpperBound <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Concat")
Private Sub Concat_AddJaggedArrayToExistingOneDimArray_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim firstArray() As Variant
    Dim secondArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    Dim expectedLength As Long
    Dim actualLength As Long
    Dim expectedUpperBound As Long
    Dim actualUpperBound As Long
    Dim testResult As Boolean
    
    firstArray = Gen.GetArray(ArrayType:=AG_ONEDIMENSION)
    secondArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    
    expected = Gen.ConcatArraysOfSameStructure(AG_JAGGED, firstArray, secondArray)
    expectedLength = Gen.GetArrayLength(expected)
    expectedUpperBound = UBound(expected)
    
    'Act:
    SUT.Items = firstArray
    actual = SUT.Concat(secondArray).Items
    actualLength = SUT.Length
    actualUpperBound = SUT.UpperBound
    testResult = SequenceEquals_JaggedArray(expected, actual)
    
    'Assert:
    Assert.IsTrue testResult, "Actual <> Expected"
    Assert.AreEqual expectedLength, actualLength, "Actual length <> expected"
    Assert.AreEqual expectedUpperBound, actualUpperBound, "Actual UpperBound <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Concat")
Private Sub Concat_AddEmptyToEmpty_ReturnsEmptyArrayWith1Slot()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    
    'Act:
    SUT.Concat expected
    ReDim expected(SUT.LowerBound)
    actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, (SUT.LowerBound = 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'''''''''''''''''''''''''''''''
' Method - CopyFromCollection '
'''''''''''''''''''''''''''''''

'TODO: CopyFromCollection test cases

'@TestMethod("BetterArray_CopyFromCollection")
Private Sub CopyFromCollection_AddCollectionToEmpty_CollectionConverted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testCollection As Collection
    Dim i As Long
    
    expected = Gen.GetArray
    Set testCollection = New Collection
    For i = LBound(expected) To UBound(expected)
        testCollection.Add expected(i)
    Next
    
    'Act:
    actual = SUT.CopyFromCollection(testCollection).Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_CopyFromCollection")
Private Sub CopyFromCollection_AddCollectionToExistingOneDimArray_ArrayReplacedWithCollectionValues()
    On Error GoTo TestFail
    
    'Arrange:
    Dim initialArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testCollection As Collection
    Dim i As Long
    
    initialArray = Gen.GetArray
    expected = Gen.GetArray
    Set testCollection = New Collection
    For i = LBound(expected) To UBound(expected)
        testCollection.Add expected(i)
    Next
    SUT.Items = initialArray
    'Act:
    actual = SUT.CopyFromCollection(testCollection).Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'''''''''''''''''''''
' Method - ToString '
'''''''''''''''''''''

'TODO: ToString test cases

'@TestMethod("BetterArray_ToString")
Private Sub ToString_FromOneDimArray_CorrectStringRepresentationReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As String = "{1,2,3}"
    Dim actual As String
    Dim testArray() As Variant
    testArray = Array(1, 2, 3)
    'Act:
    SUT.Items = testArray
    actual = SUT.ToString()

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_ToString")
Private Sub ToString_FromOneDimArrayPrettyPrint_CorrectStringRepresentationReturned()
    On Error GoTo TestFail
    
    Const expected As String = "{1, 2, 3}"
    Dim actual As String
    Dim testArray() As Variant
    testArray = Array(1, 2, 3)
    'Act:
    SUT.Items = testArray
    actual = SUT.ToString(PrettyPrint:=True)

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_ToString")
Private Sub ToString_FromOneDimArrayCustomDelimiters_CorrectStringRepresentationReturned()
    On Error GoTo TestFail
    
    Const expected As String = "[1,2,3]"
    Dim actual As String
    Dim testArray() As Variant
    testArray = Array(1, 2, 3)
    'Act:
    SUT.Items = testArray
    actual = SUT.ToString(OpeningDelimiter:="[", ClosingDelimiter:="]")
    'Assert:
    Assert.AreEqual expected, actual, "Actual <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_ToString")
Private Sub ToString_FromJaggedArray_CorrectStringRepresentationReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As String = "{{1,2},{3,4}}"
    Dim actual As String
    Dim testArray() As Variant
    
    testArray = Array(Array(1, 2), Array(3, 4))
    'Act:
    SUT.Items = testArray
    actual = SUT.ToString()

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_ToString")
Private Sub ToString_FromJaggedArrayPrettyPrint_CorrectStringRepresentationReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As String = "{" & vbCrLf _
                             & "  {1, 2}, " & vbCrLf _
                             & "  {3, 4}" & vbCrLf _
                             & "}"
    Dim actual As String
    Dim testArray() As Variant
    
    testArray = Array(Array(1, 2), Array(3, 4))
    'Act:
    SUT.Items = testArray
    actual = SUT.ToString(True)

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_ToString")
Private Sub ToString_FromEmptyArray_CorrectStringRepresentationReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As String = "{}"
    Dim actual As String
    Dim testArray() As Variant
    
    'Act:
    SUT.Items = testArray
    actual = SUT.ToString()

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_ToString")
Private Sub ToString_FromEmptyArrayPrettyPrint_CorrectStringRepresentationReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As String = "{}"
    Dim actual As String
    Dim testArray() As Variant
    
    'Act:
    SUT.Items = testArray
    actual = SUT.ToString()

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'''''''''''''''''''''
' Method - IsSorted '
'''''''''''''''''''''

'@TestMethod("BetterArray_IsSorted")
Private Sub IsSorted_SortedOneDimArray_ReturnsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Boolean
    Dim actual As Boolean
    Dim testArray() As Variant
    
    expected = True
    testArray = Array(1, 2, 3)
    SUT.Items = testArray
    'Act:
    actual = SUT.IsSorted

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_IsSorted")
Private Sub IsSorted_UnsortedOneDimArray_ReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Boolean
    Dim actual As Boolean
    Dim testArray() As Variant
    
    expected = False
    testArray = Array(2, 1, 3)
    SUT.Items = testArray
    'Act:
    actual = SUT.IsSorted

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_IsSorted")
Private Sub IsSorted_SortedMultiDimArray_ReturnsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Boolean
    Dim actual As Boolean
    Dim testArray(0 To 1, 0 To 1) As Variant
    
    expected = False
    testArray(0, 0) = "Foo"
    testArray(0, 1) = 1
    testArray(1, 0) = "Bar"
    testArray(1, 1) = 2
    SUT.Items = testArray
    'Act:
    actual = SUT.IsSorted

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_IsSorted")
Private Sub IsSorted_UnsortedMultiDimArray_ReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Boolean
    Dim actual As Boolean
    Dim testArray(0 To 1, 0 To 1) As Variant
    
    expected = False
    testArray(0, 0) = "Foo"
    testArray(0, 1) = 2
    testArray(1, 0) = "Bar"
    testArray(1, 1) = 1
    SUT.Items = testArray
    'Act:
    actual = SUT.IsSorted

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_IsSorted")
Private Sub IsSorted_SortedJaggedArray_ReturnsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Boolean
    Dim actual As Boolean
    Dim testArray() As Variant
    
    expected = True
    testArray = Array(Array("Foo", 1), Array("Bar", 1))
    SUT.Items = testArray
    'Act:
    actual = SUT.IsSorted(1)

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_IsSorted")
Private Sub IsSorted_UnsortedJaggedArray_ReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Boolean
    Dim actual As Boolean
    Dim testArray() As Variant
    
    expected = False
    testArray = Array(Array("Foo", 2), Array("Bar", 1))
    SUT.Items = testArray
    'Act:
    actual = SUT.IsSorted(1)

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_IsSorted")
Private Sub IsSorted_EmptyArray_ReturnsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Boolean
    Dim actual As Boolean
    
    expected = True
    'Act:
    actual = SUT.IsSorted

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_IsSorted")
Public Sub IsSorted_JaggedArrayWithMoreThan2Dimensions_RaisesError()
    Const ExpectedError As Long = ErrorCodes.EC_EXCEEDS_MAX_SORT_DEPTH
    On Error GoTo TestFail

    'Arrange
    Dim testArray() As Variant
    '@Ignore VariableNotUsed
    Dim actual As Boolean
    testArray = Gen.GetArray(ArrayType:=AG_JAGGED, Depth:=3)
    SUT.Items = testArray
    'Act
    actual = SUT.IsSorted

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub


'''''''''''''''''
' Method - Sort '
'''''''''''''''''

'TODO: Sort test cases

'@TestMethod("BetterArray_Sort")
Private Sub Sort_OneDimArray_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    testArray = Gen.GetArray()
    SUT.Items = testArray
    'Act:
    SUT.Sort
    actual = SUT.IsSorted
    
    'Assert:
    Assert.IsTrue actual, "Array not sorted"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Sort")
Private Sub Sort_MultiDimArray_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    testArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.Items = testArray
    'Act:
    SUT.Sort
    actual = SUT.IsSorted()
    'Assert:
    Assert.IsTrue actual, "Array not sorted"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Sort")
Private Sub Sort_JaggedArray_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    testArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.Items = testArray
    'Act:
    SUT.Sort
    actual = SUT.IsSorted()
    'Assert:
    Assert.IsTrue actual, "Array not sorted"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'''''''''''''''''''''''
' Method - CopyWithin '
'''''''''''''''''''''''

'TODO: CopyWithin test cases

'@TestMethod("BetterArray_CopyWithin")
Private Sub CopyWithin_OneDimArrayElement3ToIndex0_SelectionCopiedLengthUnchanged()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    
    testArray = Array("a", "b", "c", "d", "e")
    expected = Array("d", "b", "c", "d", "e")
    SUT.Items = testArray
    'Act:
    actual = SUT.CopyWithin(0, 3, 4).Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_CopyWithin")
Private Sub CopyWithin_OneDimArrayElements3ToEndToIndex1_SelectionCopiedLengthUnchanged()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    
    testArray = Array("a", "b", "c", "d", "e")
    expected = Array("a", "d", "e", "d", "e")
    SUT.Items = testArray
    'Act:
    actual = SUT.CopyWithin(1, 3).Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_CopyWithin")
Private Sub CopyWithin_OneDimArrayFirstTwoElementsToLastTwoElements_SelectionCopiedLengthUnchanged()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    
    testArray = Array("Banana", "Orange", "Apple", "Mango")
    expected = Array("Banana", "Orange", "Banana", "Orange")
    SUT.Items = testArray
    'Act:
    actual = SUT.CopyWithin(2, 0).Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_CopyWithin")
Private Sub CopyWithin_OneDimArrayNoStartNoEnd_NothingChanged()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    
    testArray = Array("a", "b", "c", "d", "e")
    expected = Array("a", "b", "c", "d", "e")
    SUT.Items = testArray
    'Act:
    actual = SUT.CopyWithin(0).Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_CopyWithin")
Private Sub CopyWithin_OneDimArrayPositiveStartNoEnd_SelectionCopiedLengthUnchanged()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    
    testArray = Array("a", "b", "c", "d", "e")
    expected = Array("d", "e", "c", "d", "e")
    SUT.Items = testArray
    'Act:
    actual = SUT.CopyWithin(0, 3).Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_CopyWithin")
Private Sub CopyWithin_OneDimArrayNegativeStartNoEnd_SelectionCopiedLengthUnchanged()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    
    testArray = Array("a", "b", "c", "d", "e")
    expected = Array("d", "e", "c", "d", "e")
    SUT.Items = testArray
    'Act:
    actual = SUT.CopyWithin(0, -2).Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_CopyWithin")
Private Sub CopyWithin_OneDimArrayPositiveStartPositiveEnd_SelectionCopiedLengthUnchanged()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    
    testArray = Array("a", "b", "c", "d", "e")
    expected = Array("c", "b", "c", "d", "e")
    SUT.Items = testArray
    'Act:
    actual = SUT.CopyWithin(0, 2, 3).Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_CopyWithin")
Private Sub CopyWithin_OneDimArrayPositiveStartNegativeEnd_SelectionCopiedLengthUnchanged()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    
    testArray = Array("a", "b", "c", "d", "e")
    expected = Array("c", "d", "c", "d", "e")
    SUT.Items = testArray
    'Act:
    actual = SUT.CopyWithin(0, 2, -1).Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_CopyWithin")
Private Sub CopyWithin_OneDimArrayNegativeStartNegativeEnd_SelectionCopiedLengthUnchanged()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    
    testArray = Array("a", "b", "c", "d", "e")
    expected = Array("c", "b", "c", "d", "e")
    SUT.Items = testArray
    'Act:
    actual = SUT.CopyWithin(0, -3, -2).Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_CopyWithin")
Private Sub CopyWithin_JaggedArrayElement3ToIndex0_SelectionCopiedLengthUnchanged()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testResult As Boolean
    
    testArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    expected = testArray
    expected(0) = expected(3)
    SUT.Items = testArray
    'Act:
    actual = SUT.CopyWithin(0, 3, 4).Items
    testResult = SequenceEquals_JaggedArray(expected, actual)
    'Assert:
    Assert.IsTrue testResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_CopyWithin")
Private Sub CopyWithin_EmptyInternal_RaisesError()
    Const ExpectedError As Long = ErrorCodes.EC_UNALLOCATED_ARRAY
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore VariableNotUsed
    Dim actual() As Variant
    'Act:
    actual = SUT.CopyWithin(0, 3, 4).Items
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'''''''''''''''''''
' Method - Filter '
'''''''''''''''''''

'TODO: Filter test cases

'@TestMethod("BetterArray_Filter")
Private Sub Filter_OneDimExclude_ReturnsFilteredArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    
    testArray = Array("Foo", "Bar", "Fizz", "Buzz")
    expected = Array("Foo", "Fizz", "Buzz")
    SUT.Items = testArray
    'Act:
    actual = SUT.Filter("Bar").Items

    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Filter")
Private Sub Filter_OneDimInclude_ReturnsFilteredArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    
    testArray = Array("Foo", "Bar", "Fizz", "Buzz")
    expected = Array("Bar")
    SUT.Items = testArray
    'Act:
    actual = SUT.Filter("Bar", True).Items

    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Filter")
Private Sub Filter_JaggedArrayExclude_ReturnsFilteredArray()
    On Error GoTo TestFail

    'Arrange:
    Dim testArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testResult As Boolean

    testArray = Array(Array("Foo", "Bar"), Array("Fizz", "Buzz"))
    expected = Array(Array("Foo"), Array("Fizz", "Buzz"))
    SUT.Items = testArray
    'Act:
    actual = SUT.Filter("Bar", False).Items
    testResult = SequenceEquals_JaggedArray(expected, actual)
    'Assert:
    Assert.IsTrue testResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Filter")
Private Sub Filter_JaggedArrayInclude_ReturnsFilteredArrayn()
    On Error GoTo TestFail

    'Arrange:
    Dim testArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testResult As Boolean

    testArray = Array(Array("Foo", "Bar"), Array("Fizz", "Buzz"))
    expected = Array(Array("Bar"))

    SUT.Items = testArray
    'Act:
    actual = SUT.Filter("Bar", True).Items
    testResult = SequenceEquals_JaggedArray(expected, actual)
    'Assert:
    Assert.IsTrue testResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Filter")
Private Sub Filter_MultiDimArrayExclude_ReturnsFilteredArray()
    On Error GoTo TestFail

    'Arrange:
    Dim testArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant

    ReDim testArray(1 To 2, 1 To 2)
    testArray(1, 1) = "Foo"
    testArray(1, 2) = "Bar"
    testArray(2, 1) = "Fizz"
    testArray(2, 2) = "Buzz"

    ReDim expected(1 To 2, 1 To 2)
    expected(1, 1) = "Foo"
    expected(2, 1) = "Fizz"
    expected(2, 2) = "Buzz"

    SUT.Items = testArray
    'Act:
    SUT.Filter "Bar", False
    actual = SUT.Items

    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Filter")
Private Sub Filter_MultiDimArrayInclude_ReturnsFilteredArray()
    On Error GoTo TestFail

    'Arrange:
    Dim testArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant

    ReDim testArray(1 To 2, 1 To 2)
    testArray(1, 1) = "Foo"
    testArray(1, 2) = "Bar"
    testArray(2, 1) = "Fizz"
    testArray(2, 2) = "Buzz"

    ReDim expected(1 To 1, 1 To 1)
    expected(1, 1) = "Bar"

    SUT.Items = testArray
    'Act:
    actual = SUT.Filter("Bar", True).Items
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'''''''''''''''''''''
' Method - Includes '
'''''''''''''''''''''
' TODO: add method IncludesType

'TODO: Includes test cases

'@TestMethod("BetterArray_Includes")
Private Sub Includes_OneDimArrayContainsTarget_ReturnsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim expected As Boolean
    Dim actual As Boolean
    testArray = Array("Foo", "Bar", "Fizz", "Buzz")
    expected = True
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Includes("Bar")
    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Includes")
Private Sub Includes_OneDimArrayDoesntContainTarget_ReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim expected As Boolean
    Dim actual As Boolean
    testArray = Array("Foo", "Bar", "Fizz", "Buzz")
    expected = False
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Includes("wibble")
    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Includes")
Private Sub Includes_OneDimArrayDoesContainTargetAfterStartIndex_ReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim expected As Boolean
    Dim actual As Boolean
    testArray = Array("Foo", "Bar", "Fizz", "Buzz")
    expected = True
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Includes("Fizz", 2)
    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Includes")
Private Sub Includes_OneDimArrayDoesntContainTargetAfterStartIndex_ReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim expected As Boolean
    Dim actual As Boolean
    testArray = Array("Foo", "Bar", "Fizz", "Buzz")
    expected = False
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Includes("Foo", 2)
    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Includes")
Private Sub Includes_JaggedArrayContains_ReturnsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim expected As Boolean
    Dim actual As Boolean
    testArray = Array(Array("Foo", "Bar"), Array("Fizz", "Buzz"))
    expected = True
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Includes("Buzz")
    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Includes")
Private Sub Includes_JaggedArrayDoesntContains_ReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim expected As Boolean
    Dim actual As Boolean
    testArray = Array(Array("Foo", "Bar"), Array("Fizz", "Buzz"))
    expected = False
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Includes("wibble")
    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Includes")
Private Sub Includes_EmptyInternal_ReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Boolean
    Dim actual As Boolean
    expected = False
    
    'Act:
    actual = SUT.Includes("Foo")
    
    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'''''''''''''''''
' Method - Keys '
'''''''''''''''''
'TODO: Keys test cases


'@TestMethod("BetterArray_Keys")
Private Sub Keys_OneDimArrayDefaultBase_ReturnsCorrectKeys()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim i As Long
    
    testArray = Gen.GetArray
    ReDim expected(LBound(testArray) To UBound(testArray))
    For i = LBound(testArray) To UBound(testArray)
        expected(i) = i
    Next
    SUT.Items = testArray
    'Act:
    actual = SUT.Keys
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Keys")
Private Sub Keys_OneDimArraySpecifiedBase_ReturnsCorrectKeys()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim i As Long
    
    SUT.LowerBound = 2
    testArray = Gen.GetArray
    ReDim expected(0 To Gen.GetArrayLength(testArray) - 1)
    For i = LBound(expected) To UBound(expected)
        expected(i) = i + 2
    Next
    SUT.Items = testArray
    'Act:
    actual = SUT.Keys
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Keys")
Private Sub Keys_MultiDimArray_ReturnsCorrectKeys()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim i As Long
    
    testArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    ReDim expected(LBound(testArray) To UBound(testArray))
    For i = LBound(testArray) To UBound(testArray)
        expected(i) = i
    Next
    SUT.Items = testArray
    'Act:
    actual = SUT.Keys
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Keys")
Private Sub Keys_JaggedArray_ReturnsCorrectKeys()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim i As Long
    
    testArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    ReDim expected(LBound(testArray) To UBound(testArray))
    For i = LBound(testArray) To UBound(testArray)
        expected(i) = i
    Next
    SUT.Items = testArray
    'Act:
    actual = SUT.Keys
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Keys")
Private Sub Keys_EmptyInternal_RaisesUnallocError()
    Const ExpectedError As Long = ErrorCodes.EC_UNALLOCATED_ARRAY
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore VariableNotUsed
    Dim actual() As Variant
    
    'Act:
    actual = SUT.Keys
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub



''''''''''''''''
' Method - Max '
''''''''''''''''

'TODO: Max test cases

'@TestMethod("BetterArray_Max")
Private Sub Max_OneDimArrayNumeric_ReturnsLargest()
    On Error GoTo TestFail
    
    'Arrange:


    
    'Act:

    'Assert:
    Assert.IsTrue (SUT.LowerBound = 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Max")
Private Sub Max_OneDimArrayStrings_ReturnsLargest()
    On Error GoTo TestFail
    
    'Arrange:


    
    'Act:

    'Assert:
    Assert.IsTrue (SUT.LowerBound = 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Max")
Private Sub Max_OneDimArrayVariants_ReturnsLargest()
    On Error GoTo TestFail
    
    'Arrange:


    
    'Act:

    'Assert:
    Assert.IsTrue (SUT.LowerBound = 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Max")
Private Sub Max_OneDimArrayObjects_GracefulDegradation()
    On Error GoTo TestFail
    
    'Arrange:


    
    'Act:

    'Assert:
    Assert.IsTrue (SUT.LowerBound = 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Max")
Private Sub Max_ParamArray_ReturnsLargest()
    On Error GoTo TestFail
    
    'Arrange:


    
    'Act:

    'Assert:
    Assert.IsTrue (SUT.LowerBound = 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Max")
Private Sub Max_MoreThanOneDimension_GracefulDegradation()
    On Error GoTo TestFail
    
    'Arrange:


    
    'Act:

    'Assert:
    Assert.IsTrue (SUT.LowerBound = 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Max")
Private Sub Max_EmptyInternal_GracefulDegradation()
    On Error GoTo TestFail
    
    'Arrange:


    
    'Act:

    'Assert:
    Assert.IsTrue (SUT.LowerBound = 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


''''''''''''''''
' Method - Min '
''''''''''''''''

'TODO: Min test cases

'@TestMethod("BetterArray_Min")
Private Sub Min_OneDimArrayNumeric_ReturnsSmallest()
    On Error GoTo TestFail
    
    'Arrange:


    
    'Act:

    'Assert:
    Assert.IsTrue (SUT.LowerBound = 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Min")
Private Sub Min_OneDimArrayStrings_ReturnsSmallest()
    On Error GoTo TestFail
    
    'Arrange:


    
    'Act:

    'Assert:
    Assert.IsTrue (SUT.LowerBound = 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Min")
Private Sub Min_OneDimArrayVariants_ReturnsSmallest()
    On Error GoTo TestFail
    
    'Arrange:


    
    'Act:

    'Assert:
    Assert.IsTrue (SUT.LowerBound = 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Min")
Private Sub Min_OneDimArrayObjects_GracefulDegradation()
    On Error GoTo TestFail
    
    'Arrange:


    
    'Act:

    'Assert:
    Assert.IsTrue (SUT.LowerBound = 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Min")
Private Sub Min_ParamArray_ReturnsSmallest()
    On Error GoTo TestFail
    
    'Arrange:


    
    'Act:

    'Assert:
    Assert.IsTrue (SUT.LowerBound = 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Min")
Private Sub Min_MoreThanOneDimension_GracefulDegradation()
    On Error GoTo TestFail
    
    'Arrange:


    
    'Act:

    'Assert:
    Assert.IsTrue (SUT.LowerBound = 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Min")
Private Sub Min_EmptyInternal_GracefulDegradation()
    On Error GoTo TestFail
    
    'Arrange:


    
    'Act:

    'Assert:
    Assert.IsTrue (SUT.LowerBound = 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

''''''''''''''''''
' Method - Slice '
''''''''''''''''''

'TODO: Slice test cases

'@TestMethod("BetterArray_Slice")
Private Sub Slice_OneDimNoEndArg_ReturnsCopy()
    On Error GoTo TestFail
    
    'Arrange:


    
    'Act:

    'Assert:
    Assert.IsTrue (SUT.LowerBound = 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Slice")
Private Sub Slice_OneDimWithEndArg_ReturnsCopy()
    On Error GoTo TestFail
    
    'Arrange:


    
    'Act:

    'Assert:
    Assert.IsTrue (SUT.LowerBound = 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_Slice")
Private Sub Slice_MultiDimNoEndArg_ReturnsCopy()
    On Error GoTo TestFail
    
    'Arrange:


    
    'Act:

    'Assert:
    Assert.IsTrue (SUT.LowerBound = 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Slice")
Private Sub Slice_MultiDimWithEndArg_ReturnsCopy()
    On Error GoTo TestFail
    
    'Arrange:


    
    'Act:

    'Assert:
    Assert.IsTrue (SUT.LowerBound = 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Slice")
Private Sub Slice_JaggedNoEndArg_ReturnsCopy()
    On Error GoTo TestFail
    
    'Arrange:


    
    'Act:

    'Assert:
    Assert.IsTrue (SUT.LowerBound = 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Slice")
Private Sub Slice_JaggedWithEndArg_ReturnsCopy()
    On Error GoTo TestFail
    
    'Arrange:


    
    'Act:

    'Assert:
    Assert.IsTrue (SUT.LowerBound = 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Slice")
Private Sub Slice_EmptyInternal_GracefulDegradation()
    On Error GoTo TestFail
    
    'Arrange:


    
    'Act:

    'Assert:
    Assert.IsTrue (SUT.LowerBound = 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub



''''''''''''''''''''
' Method - Reverse '
''''''''''''''''''''

'TODO: Reverse test cases

'@TestMethod("BetterArray_Reverse")
Private Sub Reverse_OneDimArray_ArrayIsReversed()
    On Error GoTo TestFail
    
    'Arrange:


    
    'Act:

    'Assert:
    Assert.IsTrue (SUT.LowerBound = 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Reverse")
Private Sub Reverse_MultiDimArray_ArrayIsReversed()
    On Error GoTo TestFail
    
    'Arrange:


    
    'Act:

    'Assert:
    Assert.IsTrue (SUT.LowerBound = 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Reverse")
Private Sub Reverse_JaggedArray_ArrayIsReversed()
    On Error GoTo TestFail
    
    'Arrange:


    
    'Act:

    'Assert:
    Assert.IsTrue (SUT.LowerBound = 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Reverse")
Private Sub Reverse_EmptyInternal_GracefulDegradation()
    On Error GoTo TestFail
    
    'Arrange:


    
    'Act:

    'Assert:
    Assert.IsTrue (SUT.LowerBound = 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub



''''''''''''''''''''
' Method - Shuffle '
''''''''''''''''''''

'TODO: Shuffle test cases

'@TestMethod("BetterArray_Shuffle")
Private Sub Shuffle_OneDimArray_ArrayIsShuffled()
    On Error GoTo TestFail
    
    'Arrange:


    
    'Act:

    'Assert:
    Assert.IsTrue (SUT.LowerBound = 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Shuffle")
Private Sub Shuffle_MultiDimArray_ArrayIsShuffled()
    On Error GoTo TestFail
    
    'Arrange:


    
    'Act:

    'Assert:
    Assert.IsTrue (SUT.LowerBound = 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_Shuffle")
Private Sub Shuffle_JaggedArray_ArrayIsShuffled()
    On Error GoTo TestFail
    
    'Arrange:


    
    'Act:

    'Assert:
    Assert.IsTrue (SUT.LowerBound = 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Shuffle")
Private Sub Shuffle_EmptyInternal_GracefulDegradation()
    On Error GoTo TestFail
    
    'Arrange:


    
    'Act:

    'Assert:
    Assert.IsTrue (SUT.LowerBound = 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'''''''''''''''''''''''''
' Method - ToExcelRange '
'''''''''''''''''''''''''

'TODO: add more test cases for ToExcelRange

'@TestMethod("BetterArray_ToExcelRange")
Private Sub ToExcelRange_OneDimensionNotTransposed_WritesValuesCorrectly()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Destination As Object
    Dim returnedRange As Object
    Dim ExcelApp As ExcelProvider
    Dim expected() As Variant
    Dim actual(TEST_ARRAY_LENGTH - 1) As Variant
    
    Set ExcelApp = New ExcelProvider
    Set Destination = ExcelApp.CurrentWorksheet.Range("A1")
    
    ' Use Array of Doubles as all values returned from an Excel range are of type Double
    expected = Gen.GetArray(AG_DOUBLE)
    SUT.Items = expected
    
    'Act:
    Set returnedRange = SUT.ToExcelRange(Destination)
    Dim i As Long
    For i = 1 To returnedRange.Columns.count
        actual(i - 1) = returnedRange.Cells.Item(1, i).Value
    Next
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Debug.Print EXCEL_DEPENDENCY_WARNING
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_ToExcelRange")
Private Sub ToExcelRange_OneDimensionTransposed_WritesValuesCorrectly()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Destination As Object
    Dim returnedRange As Object
    Dim ExcelApp As ExcelProvider
    Dim expected() As Variant
    Dim actual(TEST_ARRAY_LENGTH - 1) As Variant

    Set ExcelApp = New ExcelProvider
    Set Destination = ExcelApp.CurrentWorksheet.Range("A1")
    
    ' Use Array of Doubles as all values returned from an Excel range are of type Double
    expected = Gen.GetArray(AG_DOUBLE)
    SUT.Items = expected
    
    'Act:
    Set returnedRange = SUT.ToExcelRange(Destination, True)
    Dim i As Long
    For i = 1 To returnedRange.Rows.count
        actual(i - 1) = returnedRange.Cells.Item(i, 1).Value
    Next
    

    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Debug.Print EXCEL_DEPENDENCY_WARNING
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_ToExcelRange")
Private Sub ToExcelRange_TwoDimensionNotTransposed_WritesValuesCorrectly()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Destination As Object
    Dim ExcelApp As ExcelProvider
    Dim expected() As Variant
    Dim actual As Object
    Dim testResult As Boolean
    Dim transposed As Boolean

    Set ExcelApp = New ExcelProvider
    Set Destination = ExcelApp.CurrentWorksheet.Range("A1")
    
    expected = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.Items = expected
    transposed = False
    
    'Act:
    Set actual = SUT.ToExcelRange(Destination, transposed)
    testResult = SequenceEquals_JaggedArrayVsRange(expected, actual, transposed)
    
    'Assert:
    Assert.IsTrue testResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Debug.Print EXCEL_DEPENDENCY_WARNING
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_ToExcelRange")
Private Sub ToExcelRange_TwoDimensionTransposed_WritesValuesCorrectly()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Destination As Object
    Dim ExcelApp As ExcelProvider
    Dim expected() As Variant
    Dim actual As Object
    Dim testResult As Boolean
    Dim transposed As Boolean

    Set ExcelApp = New ExcelProvider
    Set Destination = ExcelApp.CurrentWorksheet.Range("A1")
    
    expected = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.Items = expected
    transposed = True
    
    'Act:
    Set actual = SUT.ToExcelRange(Destination, transposed)
    testResult = SequenceEquals_JaggedArrayVsRange(expected, actual, transposed)
    
    'Assert:
    Assert.IsTrue testResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Debug.Print EXCEL_DEPENDENCY_WARNING
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_ToExcelRange")
Private Sub ToExcelRange_JaggedDepthOfThree_WritesScalarRepresentationOfThirdDimension()
    On Error GoTo TestFail
    
    'Arrange:
    Dim tempBetterArray As BetterArray
    Dim Destination As Object
    Dim returnedRange As Object
    Dim OutputSheet As Object
    Dim ExcelApp As ExcelProvider
    Dim i As Long
    Dim j As Long
    Dim expected(TEST_ARRAY_LENGTH - 1, TEST_ARRAY_LENGTH - 1) As Variant
    Dim actual(TEST_ARRAY_LENGTH - 1, TEST_ARRAY_LENGTH - 1) As Variant
    Dim SourceArray() As Variant
    
    Set ExcelApp = New ExcelProvider
    Set OutputSheet = ExcelApp.CurrentWorksheet
    Set Destination = OutputSheet.Range("A1")
    
    ' Use Array of Doubles as all values returned from an Excel range are of type Double
    SourceArray = Gen.GetArray(AG_DOUBLE, AG_JAGGED, Depth:=3)
    
    For i = LBound(SourceArray) To UBound(SourceArray)
        For j = LBound(SourceArray(i)) To UBound(SourceArray(i))
            Set tempBetterArray = New BetterArray
            tempBetterArray.Items = SourceArray(i)(j)
            expected(i, j) = tempBetterArray.ToString()
            Set tempBetterArray = Nothing
        Next
    Next
    
    SUT.Items = SourceArray
    
    'Act:
    Set returnedRange = SUT.ToExcelRange(Destination)
    
    For i = 1 To returnedRange.Rows.count
        For j = 1 To returnedRange.Columns.count
            actual(i - 1, j - 1) = returnedRange.Cells.Item(i, j).Value
        Next
    Next

    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Debug.Print EXCEL_DEPENDENCY_WARNING
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub




''''''''''''''''''''''''''''
' Method - ParseFromString '
''''''''''''''''''''''''''''

'TODO: add more test cases for ParseFromString

'@TestMethod("BetterArray_ParseFromString")
Private Sub ParseFromString_OneDimensionArrayFromToString_ReturnsCorrectValues()
    On Error GoTo TestFail
    
    'Arrange:
    Dim tempBetterArray As BetterArray
    Dim expected() As Variant
    Dim actual() As Variant
    Dim SourceString As String
    Dim testResult As Boolean

    Set tempBetterArray = New BetterArray
    expected = Gen.GetArray()
    tempBetterArray.Items = expected
    SourceString = tempBetterArray.ToString()
    
    'Act:
    actual = SUT.ParseFromString(SourceString).Items
    
    ' can't use Assert.SequenceEquals due to type comparison - Bytes Will be Long in actual
    testResult = SequenceEquals_JaggedArray(expected, actual)
    
    'Assert:
    Assert.IsTrue testResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_ParseFromString")
Private Sub ParseFromString_Jagged2DeepArrayFromToString_ReturnsCorrectValues()
    On Error GoTo TestFail
    
    'Arrange:
    Dim tempBetterArray As BetterArray
    Dim expected() As Variant
    Dim actual() As Variant
    Dim SourceString As String
    Dim testResult As Boolean
    
    Set tempBetterArray = New BetterArray
    
    expected = Gen.GetArray(AG_BYTE, AG_JAGGED)
    tempBetterArray.Items = expected
    SourceString = tempBetterArray.ToString()
    
    'Act:
    actual = SUT.ParseFromString(SourceString).Items
    
    ' can't use Assert.SequenceEquals due to type comparison - Bytes Will be Long in actual
    testResult = SequenceEquals_JaggedArray(expected, actual)
    
    'Assert:
    Assert.IsTrue testResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_ParseFromString")
Private Sub ParseFromString_Jagged3DeepArrayFromToString_ReturnsCorrectValues()
    On Error GoTo TestFail
    
    'Arrange:
    Dim tempBetterArray As BetterArray
    Dim expected() As Variant
    Dim actual() As Variant
    Dim SourceString As String
    Dim testResult As Boolean
    
    Set tempBetterArray = New BetterArray
    expected = Gen.GetArray(AG_BYTE, AG_JAGGED, Depth:=3)
    tempBetterArray.Items = expected
    SourceString = tempBetterArray.ToString()
    
    'Act:
    actual = SUT.ParseFromString(SourceString).Items
    
    ' can't use Assert.SequenceEquals due to type comparison - Bytes Will be Long in actual
    ' also, Assert.SeqenceEquals doesn't support jagged arrays: https://github.com/rubberduck-vba/Rubberduck/issues/5161
    testResult = SequenceEquals_JaggedArray(expected, actual)
    
    
    'Assert:
    Assert.IsTrue testResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_ParseFromString")
Private Sub ParseFromString_Jagged5DeepArrayFromToString_ReturnsCorrectValues()
    On Error GoTo TestFail
    
    'Arrange:
    Dim tempBetterArray As BetterArray
    Dim expected As String
    Dim actual As String
    
    Set tempBetterArray = New BetterArray
    tempBetterArray.Items = Gen.GetArray(AG_BYTE, AG_JAGGED, 5)
    expected = tempBetterArray.ToString()
    
    'Act:
    actual = SUT.ParseFromString(expected).ToString
        
    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub








