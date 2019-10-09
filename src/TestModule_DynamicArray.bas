Attribute VB_Name = "TestModule_DynamicArray"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

'@IgnoreModule ProcedureNotUsed
'@IgnoreModule LineLabelNotUsed
'@IgnoreModule EmptyMethod

Private Assert As Object
'Move to early bind
' Private Assert As AssertClass
'@Ignore VariableNotUsed
Private Fakes As Object
'Move to early bind
'Private Fakes As FakesProvider


Private Const TEST_ARRAY_LENGTH As Long = 10

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    ' Move to early binding
'    Set Assert = New AssertClass
'    Set Fakes = New FakesProvider
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
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

''@TestMethod("Uncategorized")
'Private Sub MethodToTest_Scenario_ExpectedBehaviour()                        'Example
'    On Error GoTo TestFail
'
'    'Arrange:
'    'Act:
'
'    'Assert:
'    Assert.Succeed
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'End Sub


'''''''''''''''''
' Instantiation '
'''''''''''''''''

'@TestMethod("DynamicArrayConstructor")
Private Sub DynamicArray_CanInstantiate_SUTNotNothing()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    'Act:
    Set SUT = New DynamicArray
    'Assert:
    Assert.IsNotNothing SUT

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DynamicArrayConstructor")
Private Sub DynamicArray_CreatesWithDefaultCapacity_CapacityIsFour()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As Long = 4
    Dim SUT As DynamicArray
    Dim actual As Long
    
    Set SUT = New DynamicArray
    
    'Act:
    actual = SUT.Capacity
    
    'Assert:
    Assert.AreEqual expected, actual, "Default capacity incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'''''''''''''''''''
' Prop - Capacity '
'''''''''''''''''''

'@TestMethod("DynamicArray_Capacity")
Private Sub Capacity_CanSetCapacity_ReturnedCapacityMatchesSetCapacity()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As Long = 20
    Dim SUT As DynamicArray
    Dim actual As Long

    Set SUT = New DynamicArray
    
    'Act:
    SUT.Capacity = expected
    actual = SUT.Capacity

    'Assert:
    Assert.AreEqual expected, actual, "Returned capacity does not equal set capacity"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


''''''''''''''''
' Prop - Items '
''''''''''''''''

'@TestMethod("DynamicArray_Items")
Private Sub Items_CanAssignOneDimemsionalArray_ReturnedArrayEqualsAssignedArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim gen As ArrayGenerator
    Dim SUT As DynamicArray
    Dim expected As Variant
    Dim actual As Variant
    
    Set gen = New ArrayGenerator
    expected = gen.getArray(10, VariantVals, OneDimension)
    Set SUT = New DynamicArray
    
    'Act:
    SUT.Items = expected
    actual = SUT.Items

    'Assert:
    Assert.SequenceEquals expected, actual, "Actual array does not match expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Items")
Private Sub Items_CanAssignMultiDimemsionalArray_ReturnedArrayEqualsAssignedArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim gen As ArrayGenerator
    Dim SUT As DynamicArray
    Dim expected As Variant
    Dim actual As Variant
    
    Set gen = New ArrayGenerator
    expected = gen.getArray(10, VariantVals, MultiDimension)
    Set SUT = New DynamicArray
    
    'Act:
    SUT.Items = expected
    actual = SUT.Items

    'Assert:
    Assert.SequenceEquals expected, actual, "Actual array does not match expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Items")
' NOTE: does not use SequenceEquals due to Rubberduck issue: https://github.com/rubberduck-vba/Rubberduck/issues/5161
Private Sub Items_CanAssignJaggedArray_ReturnedArrayEqualsAssignedArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim gen As ArrayGenerator
    Dim SUT As DynamicArray
    Dim expected As Variant
    Dim actual As Variant
    Dim i As Long
    Dim j As Long
    Dim testResult As Boolean
    
    Set gen = New ArrayGenerator
    expected = gen.getArray(TEST_ARRAY_LENGTH, VariantVals, Jagged)
    Set SUT = New DynamicArray
    
    'Act:
    SUT.Items = expected
    actual = SUT.Items
    testResult = True
    For i = LBound(actual) To UBound(actual)
        For j = LBound(actual(i)) To UBound(actual(i))
            If actual(i)(j) <> expected(i)(j) Then
                testResult = False
                Exit For
            End If
        Next
        If testResult = False Then Exit For
    Next

    'Assert:
    Assert.IsTrue testResult, "Contents of expected and actual do not match"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'''''''''''''''''
' Prop - Length '
'''''''''''''''''

'@TestMethod("DynamicArray_Length")
Private Sub Length_FromAssignedOneDimensionalArray_ReturnedLengthEqualsOriginalArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim gen As ArrayGenerator
    Dim SUT As DynamicArray
    Dim testArray As Variant
    Dim expected As Long
    Dim actual As Long
    
    Set gen = New ArrayGenerator
    expected = TEST_ARRAY_LENGTH
    testArray = gen.getArray(TEST_ARRAY_LENGTH, VariantVals, OneDimension)
    Set SUT = New DynamicArray
    
    'Act:
    SUT.Items = testArray
    actual = SUT.Length

    'Assert:
    Assert.AreEqual expected, actual, "Actual length <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Length")
Private Sub Length_FromAssignedMultiDimensionalArray_ReturnedLengthEqualsOriginalArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim gen As ArrayGenerator
    Dim SUT As DynamicArray
    Dim testArray As Variant
    Dim expected As Long
    Dim actual As Long
    
    Set gen = New ArrayGenerator
    expected = TEST_ARRAY_LENGTH
    testArray = gen.getArray(TEST_ARRAY_LENGTH, VariantVals, MultiDimension)
    Set SUT = New DynamicArray
    
    'Act:
    SUT.Items = testArray
    actual = SUT.Length
    
    'Assert:
    Assert.AreEqual expected, actual, "Actual length <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Length")
Private Sub Length_FromAssignedJaggedArray_ReturnedLengthEqualsOriginalArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim gen As ArrayGenerator
    Dim SUT As DynamicArray
    Dim testArray As Variant
    Dim expected As Long
    Dim actual As Long
    
    
    Set gen = New ArrayGenerator
    expected = TEST_ARRAY_LENGTH
    testArray = gen.getArray(TEST_ARRAY_LENGTH, VariantVals, Jagged)
    Set SUT = New DynamicArray
    
    'Act:
    SUT.Items = testArray
    actual = SUT.Length

    'Assert:
    Assert.AreEqual expected, actual, "Actual length <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DynamicArray_Length")
Private Sub Upperbound_FromAssignedOneDimensionalArray_ReturnedUpperBoundEqualsOriginalArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim gen As ArrayGenerator
    Dim SUT As DynamicArray
    Dim testArray As Variant
    Dim expected As Long
    Dim actual As Long
    
    Set gen = New ArrayGenerator
    testArray = gen.getArray(TEST_ARRAY_LENGTH, VariantVals, OneDimension)
    expected = UBound(testArray)
    Set SUT = New DynamicArray
    
    'Act:
    SUT.Items = testArray
    actual = SUT.UpperBound

    'Assert:
    Assert.AreEqual expected, actual, "Actual upperbound <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'''''''''''''''
' Prop - Base '
'''''''''''''''


'@TestMethod("DynamicArray_Base")
Private Sub Base_FromAssignedOneDimensionalArray_ReturnedBaseEqualsOriginalArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim gen As ArrayGenerator
    Dim SUT As DynamicArray
    Dim testArray As Variant
    Dim expected As Long
    Dim actual As Long
    
    Set gen = New ArrayGenerator
    testArray = gen.getArray(TEST_ARRAY_LENGTH, VariantVals, OneDimension)
    expected = LBound(testArray)
    Set SUT = New DynamicArray
    
    'Act:
    SUT.Items = testArray
    actual = SUT.Base

    'Assert:
    Assert.AreEqual expected, actual, "Actual base <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DynamicArray_Base")
Private Sub Base_ChangingBaseOfAssignedArray_ReturnedArrayHasNewBase()
    On Error GoTo TestFail
    
    'Arrange:
    Dim gen As ArrayGenerator
    Dim SUT As DynamicArray
    Dim testArray As Variant
    Dim returnedItems As Variant
    Dim expected As Long
    Dim actual As Long
    Dim oldBase As Long
    
    Set gen = New ArrayGenerator
    testArray = gen.getArray(TEST_ARRAY_LENGTH, VariantVals, OneDimension)
    oldBase = LBound(testArray)
    Set SUT = New DynamicArray
    
    'Act:
    SUT.Items = testArray
    expected = oldBase + 1
    SUT.Base = expected
    returnedItems = SUT.Items
    actual = LBound(returnedItems)

    'Assert:
    Assert.AreEqual expected, actual, "Actual base <> expected"
    Assert.AreEqual SUT.Base, actual, "Actual base <> SUT.Base prop"
    Assert.AreEqual UBound(testArray) + 1, UBound(returnedItems), "Actual upperbound <> expected"
    Assert.AreEqual SUT.UpperBound, UBound(returnedItems), "Actual upperbound <> SUT.UpperBound prop"
    Assert.AreEqual TEST_ARRAY_LENGTH, SUT.Length, "Actual length does not equal expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'''''''''''''''
' Prop - Item '
'''''''''''''''

'@TestMethod("DynamicArray_Item")
Private Sub Item_ChangingExistingIndex_ItemIsChanged()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As String = "Hello World"
    Dim gen As ArrayGenerator
    Dim SUT As DynamicArray
    Dim testArray As Variant
    Dim actual As Variant
    Dim actualBase As Long
    Dim expectedBase As Long

    
    Set gen = New ArrayGenerator
    testArray = gen.getArray(TEST_ARRAY_LENGTH, VariantVals, OneDimension)
    expectedBase = LBound(testArray)
    Set SUT = New DynamicArray
    
    'Act:
    SUT.Items = testArray
    SUT.Item(1) = expected
    actual = SUT.Item(1)
    actualBase = SUT.Base

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
    Assert.AreEqual expectedBase, actualBase, "Actual base does not equal expected base"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Item")
Private Sub Item_ChangingIndexOverUpperBound_ItemIsPushed()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As String = "Hello World"
    Dim gen As ArrayGenerator
    Dim SUT As DynamicArray
    Dim testArray As Variant
    Dim actual As Variant
    
    Dim actualBase As Long
    Dim expectedBase As Long
    
    Set gen = New ArrayGenerator
    testArray = gen.getArray(TEST_ARRAY_LENGTH, VariantVals, OneDimension)
    expectedBase = LBound(testArray)
    Set SUT = New DynamicArray
    
    'Act:
    SUT.Items = testArray
    SUT.Item(SUT.UpperBound + 1) = expected
    actual = SUT.Item(SUT.UpperBound)
    actualBase = SUT.Base
    

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
    Assert.AreEqual TEST_ARRAY_LENGTH + 1, SUT.Length, "Actual length does not match expected length"
    Assert.AreEqual expectedBase, actualBase, "Actual base does not match expected base"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Item")
Private Sub Item_ChangingIndexBelowBase_ItemIsUnshifted()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As String = "Hello World"
    Dim gen As ArrayGenerator
    Dim SUT As DynamicArray
    Dim testArray As Variant
    Dim actual As Variant
    
    Dim expectedBase As Long
    Dim actualBase As Long

    
    Set gen = New ArrayGenerator
    testArray = gen.getArray(TEST_ARRAY_LENGTH, VariantVals, OneDimension)
    expectedBase = LBound(testArray)
    Set SUT = New DynamicArray
    
    'Act:
    SUT.Items = testArray
    SUT.Item(SUT.Base - 10) = expected
    actual = SUT.Item(SUT.Base)
    actualBase = SUT.Base

    'Assert:
    Assert.AreEqual expected, actual, "Actual result does not match expected result"
    Assert.AreEqual TEST_ARRAY_LENGTH + 1, SUT.Length, "Actual length does not match expected length"
    Assert.AreEqual expectedBase, actualBase, "Actual base does not match expected base"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DynamicArray_Item")
Private Sub Item_GetScalarValue_ValueReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim gen As ArrayGenerator
    Dim SUT As DynamicArray
    Dim testArray As Variant
    Dim expected As Variant
    Dim actual As Variant
       
    Set gen = New ArrayGenerator
    Set SUT = New DynamicArray
    testArray = gen.getArray(TEST_ARRAY_LENGTH, VariantVals, OneDimension)
    expected = testArray(1)
    
    'Act:
    SUT.Items = testArray
    actual = SUT.Item(1)

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DynamicArray_Item")
Private Sub Item_GetObject_SameObjectReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim gen As ArrayGenerator
    Dim SUT As DynamicArray
    Dim testArray As Variant
    Dim expected As Object
    Dim actual As Object
       
    Set gen = New ArrayGenerator
    testArray = gen.getArray(TEST_ARRAY_LENGTH, ObjectVals, OneDimension)
    Set SUT = New DynamicArray
    Set expected = testArray(1)
    
    'Act:
    SUT.Items = testArray
    Set actual = SUT.Item(1)

    'Assert:
    Assert.AreSame expected, actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'''''''''''''''''
' Method - Push '
'''''''''''''''''

'@TestMethod("DynamicArray_Push")
Private Sub Push_AddToNewDynamicArray_ItemAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As String = "Hello World"
    Const expectedLength As Long = 1
    Const expectedUpperBound As Long = 0
    Dim SUT As DynamicArray
    Dim actual As String
    Dim actualLength As Long
    Dim actualUpperBound As Long
    
    Set SUT = New DynamicArray
    
    'Act:
    SUT.Push expected
    actual = SUT.Item(SUT.Base)
    actualLength = SUT.Length
    actualUpperBound = SUT.UpperBound

    'Assert:
    Assert.AreEqual expected, actual, "actual <> expected"
    Assert.AreEqual expectedLength, actualLength, "Length value incorrect"
    Assert.AreEqual expectedUpperBound, actualUpperBound, "Upperbound value incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Push")
Private Sub Push_AddToExistingOneDimensionalArray_ItemAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As String = "Hello World"
    Dim SUT As DynamicArray
    Dim testArray As Variant
    Dim gen As ArrayGenerator
    Dim actual As String
    
    Set SUT = New DynamicArray
    Set gen = New ArrayGenerator
    testArray = gen.getArray(TEST_ARRAY_LENGTH, VariantVals, OneDimension)
    
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
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Push")
Private Sub Push_AddToExistingMultidimensionalArray_ItemAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Dim expected As Variant
    Dim actual As Variant

    Dim testArray As Variant
    Dim returnedArray As Variant
    Dim gen As ArrayGenerator
    
    Set SUT = New DynamicArray
    Set gen = New ArrayGenerator
    
    expected = "Hello World"
    testArray = gen.getArray(TEST_ARRAY_LENGTH, VariantVals, MultiDimension)
    
    'Act:
    SUT.Items = testArray
    SUT.Push expected
    returnedArray = SUT.Items
    actual = returnedArray(UBound(returnedArray), UBound(returnedArray, 2))

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
    Assert.AreEqual TEST_ARRAY_LENGTH + 1, SUT.Length, "Length value incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DynamicArray_Push")
Private Sub Push_AddToExistingJaggedArray_ItemAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Dim expected As Variant
    Dim actual As Variant

    Dim testArray As Variant
    Dim returnedArray As Variant
    Dim gen As ArrayGenerator
    
    Set SUT = New DynamicArray
    Set gen = New ArrayGenerator
    
    expected = gen.getArray(TEST_ARRAY_LENGTH, VariantVals, OneDimension)
    testArray = gen.getArray(TEST_ARRAY_LENGTH, VariantVals, Jagged)
    
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
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DynamicArray_Push")
Private Sub Push_AddMultipleToNewDynamicArray_ItemsAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As Long = 1
    Const expectedLength As Long = 3
    Const expectedUpperBound As Long = 2
    Dim SUT As DynamicArray
    Dim actual As Long
    Dim actualLength As Long
    Dim actualUpperBound As Long
    
    Set SUT = New DynamicArray
    
    
    'Act:
    SUT.Push expected, 2, 3
    actual = SUT.Item(SUT.Base)
    actualLength = SUT.Length
    actualUpperBound = SUT.UpperBound

    'Assert:
    Assert.AreEqual expected, actual, "Element value incorrect"
    Assert.AreEqual expectedLength, actualLength, "Length value incorrect"
    Assert.AreEqual expectedUpperBound, actualUpperBound, "Upperbound value incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


''''''''''''''''
' Method - Pop '
''''''''''''''''

'@TestMethod("DynamicArray_Pop")
Private Sub Pop_OneDimensionalArray_LastItemRemoved()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    Dim gen As ArrayGenerator
    Set gen = New ArrayGenerator
    Dim testArray As Variant
    Dim actualBase As Long
    Dim expectedBase As Long
    
    testArray = gen.getArray(TEST_ARRAY_LENGTH, StringVals, OneDimension)
    Dim expected As String
    Dim actual As String

    expected = testArray(UBound(testArray))
    expectedBase = SUT.Base
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Pop
    actualBase = SUT.Base

    'Assert:
    Assert.AreEqual expected, actual, "Element value incorrect"
    Assert.AreEqual TEST_ARRAY_LENGTH - 1, SUT.Length, "Length value incorrect"
    Assert.AreEqual UBound(testArray) - 1, SUT.UpperBound, "Upperbound value incorrect"
    Assert.AreEqual expectedBase, actualBase, "Base value incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DynamicArray_Pop")
Private Sub Pop_ArrayLengthIsZero_ReturnsEmpty()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    Dim expected As Variant
    Dim expectedBase As Long
    Dim expectedLength As Long
    Dim expectedUpperBound As Long
    Dim actual As Variant
    Dim actualBase As Long
    Dim actualLength As Long
    Dim actualUpperBound As Long
    
    expected = Empty
    expectedBase = 0
    expectedLength = 0
    expectedUpperBound = 0
    
    'Act:
    actual = SUT.Pop
    actualBase = SUT.Base
    actualLength = SUT.Length
    actualUpperBound = SUT.UpperBound

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
    Assert.AreEqual expectedBase, actualBase, "Base value incorrect"
    Assert.AreEqual expectedLength, actualLength, "Length value incorrect"
    Assert.AreEqual expectedUpperBound, actualUpperBound, "Upperbound value incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


''''''''''''''''''
' Method - Shift '
''''''''''''''''''

'@TestMethod("DynamicArray_Shift")
Private Sub Shift_OneDimensionalArray_FirstItemRemoved()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    Dim gen As ArrayGenerator
    Set gen = New ArrayGenerator
    Dim testArray As Variant
    Dim actualBase As Long
    Dim expectedBase As Long
    
    testArray = gen.getArray(TEST_ARRAY_LENGTH, StringVals, OneDimension)
    Dim expected As String
    Dim actual As String

    expected = testArray(LBound(testArray))
    expectedBase = SUT.Base
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Shift
    actualBase = SUT.Base

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
    Assert.AreEqual TEST_ARRAY_LENGTH - 1, SUT.Length, "Length value incorrect"
    Assert.AreEqual UBound(testArray) - 1, SUT.UpperBound, "Upperbound value incorrect"
    Assert.AreEqual expectedBase, actualBase, "Base value incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Shift")
Private Sub Shift_ArrayLengthIsZero_ReturnsEmpty()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As Variant = Empty
    Const expectedBase As Long = 0
    Const expectedLength As Long = 0
    Const expectedUpperBound As Long = 0
    Dim SUT As DynamicArray
    Dim actual As Variant
    Dim actualBase As Long
    Dim actualLength As Long
    Dim actualUpperBound As Long
    
    Set SUT = New DynamicArray
    
    'Act:
    actual = SUT.Shift
    actualBase = SUT.Base
    actualLength = SUT.Length
    actualUpperBound = SUT.UpperBound

    'Assert:
    Assert.AreEqual expected, actual, "Element value incorrect"
    Assert.AreEqual expectedBase, actualBase, "Base value incorrect"
    Assert.AreEqual expectedLength, actualLength, "Length value incorrect"
    Assert.AreEqual expectedUpperBound, actualUpperBound, "Upperbound value incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

''''''''''''''''''''
' Method - Unshift '
''''''''''''''''''''

'@TestMethod("DynamicArray_Unshift")
Private Sub Unshift_OneDimensionalArray_ItemAddedToBeginning()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Dim gen As ArrayGenerator
    Dim testArray As Variant
    Dim expected As String
    Dim actual As String
    Dim actualBase As Long
    Dim expectedBase As Long
    Dim testElement As String
    
    Set SUT = New DynamicArray
    Set gen = New ArrayGenerator
    testArray = gen.getArray(TEST_ARRAY_LENGTH, StringVals, OneDimension)
    testElement = "Hello World"
    expectedBase = SUT.Base
    expected = TEST_ARRAY_LENGTH + 1
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Unshift(testElement)
    actualBase = SUT.Base

    'Assert:
    Assert.AreEqual expected, actual, "Return value incorrect"
    Assert.AreEqual (UBound(testArray) + 1), SUT.UpperBound, "Upperbound value incorrect"
    Assert.AreEqual expectedBase, actualBase, "Base value incorrect"
    Assert.AreEqual testElement, SUT.Item(SUT.Base), "Element not inserted at correct position"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DynamicArray_Unshift")
Private Sub Unshift_ArrayLengthIsZero_ItemIsPushedToEmptyArray()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As Long = 1
    Const expectedBase As Long = 0
    Const expectedUpperBound As Long = 0
    Const expectedElement As String = "Hello World"
    Dim SUT As DynamicArray
    Dim actual As Long
    Dim actualBase As Long
    Dim actualUpperBound As Long
    Dim actualElement As String

    Set SUT = New DynamicArray
    
    'Act:
    actual = SUT.Unshift(expectedElement)
    actualBase = SUT.Base
    actualUpperBound = SUT.UpperBound
    actualElement = SUT.Item(SUT.Base)

    'Assert:
    Assert.AreEqual expected, actual, "Actual length <> expected length"
    Assert.AreEqual expectedUpperBound, actualUpperBound, "Upperbound value incorrect"
    Assert.AreEqual expectedBase, actualBase, "Base value incorrect"
    Assert.AreEqual expectedElement, actualElement, "Actual element <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DynamicArray_Unshift")
Public Sub Unshift_MultidimensionalArray_ItemAddedToBeginning()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As Long = TEST_ARRAY_LENGTH + 1
    Const expectedBase As Long = 0
    Const expectedUpperBound As Long = TEST_ARRAY_LENGTH
    Const expectedElement As String = "Hello World"
    Dim SUT As DynamicArray
    Dim gen As ArrayGenerator
    Dim actual As Long
    Dim actualBase As Long
    Dim actualUpperBound As Long
    Dim actualElement As String
    Dim testArray As Variant
    Dim returnedItems As Variant

    Set SUT = New DynamicArray
    Set gen = New ArrayGenerator
    testArray = gen.getArray(TEST_ARRAY_LENGTH, VariantVals, MultiDimension)
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Unshift(expectedElement)
    returnedItems = SUT.Items
    actualBase = SUT.Base
    actualUpperBound = SUT.UpperBound
    actualElement = returnedItems(LBound(returnedItems), LBound(returnedItems, 2))

    'Assert:
    Assert.AreEqual expected, actual, "Actual length <> expected length"
    Assert.AreEqual expectedUpperBound, actualUpperBound, "Upperbound value incorrect"
    Assert.AreEqual expectedBase, actualBase, "Base value incorrect"
    Assert.AreEqual expectedElement, actualElement, "Actual element <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'''''''''''''''''''
' Method - Concat '
'''''''''''''''''''
'TODO: Concat test cases

'@TestMethod("DynamicArray_Concat")
Public Sub Concat_AddOneDimArrayToEmptyInternal_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Concat")
Public Sub Concat_AddOneDimArrayToExistingOneDimArray_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Concat")
Public Sub Concat_AddMultiDimArrayToEmptyInternal_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Concat")
Public Sub Concat_AddMultiDimArrayToExistingMultiDimArray_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DynamicArray_Concat")
Public Sub Concat_AddJaggedArrayToEmptyInternal_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Concat")
Public Sub Concat_AddJaggedArrayToExistingJagged_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Concat")
Public Sub Concat_AddOneDimArrayToExistingJagged_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Concat")
Public Sub Concat_AddOneDimArrayToExistingMulti_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Concat")
Public Sub Concat_AddMultiDimArrayToExistingOneDimArray_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Concat")
Public Sub Concat_AddJaggedArrayToExistingOneDimArray_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Concat")
Public Sub Concat_AddEmptyToEmpty_GracefulDegradation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'''''''''''''''''''''''''''''''
' Method - CopyFromCollection '
'''''''''''''''''''''''''''''''

'TODO: CopyFromCollection test cases

'@TestMethod("DynamicArray_CopyFromCollection")
Public Sub CopyFromCollection_AddCollectionToEmpty_CollectionConverted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_CopyFromCollection")
Public Sub CopyFromCollection_AddCollectionToExistingOneDimArray_ArrayReplacedWithCollectionValues()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'''''''''''''''''''''
' Method - ToString '
'''''''''''''''''''''

'TODO: ToString test cases

'@TestMethod("DynamicArray_ToString")
Public Sub ToString_FromOneDimArray_CorrectStringRepresentationReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_ToString")
Public Sub ToString_FromOneDimArrayPrettyPrint_CorrectStringRepresentationReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_ToString")
Public Sub ToString_FromJaggedArray_CorrectStringRepresentationReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_ToString")
Public Sub ToString_FromJaggedArrayPrettyPrint_CorrectStringRepresentationReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DynamicArray_ToString")
Public Sub ToString_FromEmptyArray_CorrectStringRepresentationReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_ToString")
Public Sub ToString_FromEmptyArrayPrettyPrint_CorrectStringRepresentationReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'''''''''''''''''''''
' Method - ToString '
'''''''''''''''''''''

'TODO: Sort test cases

'@TestMethod("DynamicArray_Sort")
Public Sub Sort_OneDimArray_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Sort")
Public Sub Sort_MultiDimArray_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Sort")
Public Sub Sort_JaggedArray_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'''''''''''''''''''''''
' Method - CopyWithin '
'''''''''''''''''''''''

'TODO: CopyWithin test cases

'@TestMethod("DynamicArray_CopyWithin")
Public Sub CopyWithin_OneDimArrayNoStartNoEnd_SelectionShallowCopiedLengthUnchanged()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_CopyWithin")
Public Sub CopyWithin_OneDimArrayPositiveStartNoEnd_SelectionShallowCopiedLengthUnchanged()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_CopyWithin")
Public Sub CopyWithin_OneDimArrayNegativeStartNoEnd_SelectionShallowCopiedLengthUnchanged()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_CopyWithin")
Public Sub CopyWithin_OneDimArrayPositiveStartPositiveEnd_SelectionShallowCopiedLengthUnchanged()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_CopyWithin")
Public Sub CopyWithin_OneDimArrayPositiveStartNegativeEnd_SelectionShallowCopiedLengthUnchanged()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_CopyWithin")
Public Sub CopyWithin_OneDimArrayNegativeStartNegativeEnd_SelectionShallowCopiedLengthUnchanged()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_CopyWithin")
Public Sub CopyWithin_JaggedArrayNoStartNoEnd_SelectionShallowCopiedLengthUnchanged()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_CopyWithin")
Public Sub CopyWithin_JaggedArrayPositiveStartNoEnd_SelectionShallowCopiedLengthUnchanged()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_CopyWithin")
Public Sub CopyWithin_JaggedArrayNegativeStartNoEnd_SelectionShallowCopiedLengthUnchanged()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_CopyWithin")
Public Sub CopyWithin_JaggedArrayPositiveStartPositiveEnd_SelectionShallowCopiedLengthUnchanged()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_CopyWithin")
Public Sub CopyWithin_JaggedArrayPositiveStartNegativeEnd_SelectionShallowCopiedLengthUnchanged()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_CopyWithin")
Public Sub CopyWithin_JaggedArrayNegativeStartNegativeEnd_SelectionShallowCopiedLengthUnchanged()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_CopyWithin")
Public Sub CopyWithin_EmptyInternal_GracefulDegradation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'''''''''''''''''''
' Method - Filter '
'''''''''''''''''''

'TODO: Filter test cases

'@TestMethod("DynamicArray_Filter")
Public Sub Filter_OneDimExclude_ReturnsFilteredArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Filter")
Public Sub Filter_OneDimInclude_ReturnsFilteredArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Filter")
Public Sub Filter_ArrayMoreThanOneDimension_GracefulDegradation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Filter")
Public Sub Filter_EmptyInternal_GracefulDegradation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'''''''''''''''''''''
' Method - Includes '
'''''''''''''''''''''

'TODO: Includes test cases

'@TestMethod("DynamicArray_Includes")
Public Sub Includes_OneDimArrayContainsTarget_ReturnsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Includes")
Public Sub Includes_OneDimArrayDoesNotContainTarget_ReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Includes")
Public Sub Includes_ArrayMoreThanOneDimension_GracefulDegradation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Includes")
Public Sub Includes_EmptyInternal_GracefulDegradation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'''''''''''''''''
' Method - Keys '
'''''''''''''''''
'TODO: Keys test cases


'@TestMethod("DynamicArray_Keys")
Public Sub Keys_OneDimArray_ReturnsCorrectKeys()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Keys")
Public Sub Keys_MultiDimArray_ReturnsCorrectKeys()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Keys")
Public Sub Keys_JaggedArray_ReturnsCorrectKeys()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Keys")
Public Sub Keys_EmptyInternal_ReturnsCorrectKeys()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


''''''''''''''''
' Method - Max '
''''''''''''''''

'TODO: Max test cases

'@TestMethod("DynamicArray_Max")
Public Sub Max_OneDimArrayNumeric_ReturnsLargest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Max")
Public Sub Max_OneDimArrayStrings_ReturnsLargest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Max")
Public Sub Max_OneDimArrayVariants_ReturnsLargest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Max")
Public Sub Max_OneDimArrayObjects_GracefulDegradation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Max")
Public Sub Max_ParamArray_ReturnsLargest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Max")
Public Sub Max_MoreThanOneDimension_GracefulDegradation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Max")
Public Sub Max_EmptyInternal_GracefulDegradation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


''''''''''''''''
' Method - Min '
''''''''''''''''

'TODO: Min test cases

'@TestMethod("DynamicArray_Min")
Public Sub Min_OneDimArrayNumeric_ReturnsSmallest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Min")
Public Sub Min_OneDimArrayStrings_ReturnsSmallest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Min")
Public Sub Min_OneDimArrayVariants_ReturnsSmallest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Min")
Public Sub Min_OneDimArrayObjects_GracefulDegradation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Min")
Public Sub Min_ParamArray_ReturnsSmallest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Min")
Public Sub Min_MoreThanOneDimension_GracefulDegradation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Min")
Public Sub Min_EmptyInternal_GracefulDegradation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

''''''''''''''''''
' Method - Slice '
''''''''''''''''''

'TODO: Slice test cases

'@TestMethod("DynamicArray_Slice")
Public Sub Slice_OneDimNoEndArg_ReturnsShallowCopy()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Slice")
Public Sub Slice_OneDimWithEndArg_ReturnsShallowCopy()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DynamicArray_Slice")
Public Sub Slice_MultiDimNoEndArg_ReturnsShallowCopy()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Slice")
Public Sub Slice_MultiDimWithEndArg_ReturnsShallowCopy()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Slice")
Public Sub Slice_JaggedNoEndArg_ReturnsShallowCopy()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Slice")
Public Sub Slice_JaggedWithEndArg_ReturnsShallowCopy()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Slice")
Public Sub Slice_EmptyInternal_GracefulDegradation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub



''''''''''''''''''''
' Method - Reverse '
''''''''''''''''''''

'TODO: Reverse test cases

'@TestMethod("DynamicArray_Reverse")
Public Sub Reverse_OneDimArray_ArrayIsReversed()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Reverse")
Public Sub Reverse_MultiDimArray_ArrayIsReversed()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Reverse")
Public Sub Reverse_JaggedArray_ArrayIsReversed()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Reverse")
Public Sub Reverse_EmptyInternal_GracefulDegradation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub



''''''''''''''''''''
' Method - Shuffle '
''''''''''''''''''''

'TODO: Shuffle test cases

'@TestMethod("DynamicArray_Shuffle")
Public Sub Shuffle_OneDimArray_ArrayIsShuffled()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Shuffle")
Public Sub Shuffle_MultiDimArray_ArrayIsShuffled()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DynamicArray_Shuffle")
Public Sub Shuffle_JaggedArray_ArrayIsShuffled()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Shuffle")
Public Sub Shuffle_EmptyInternal_GracefulDegradation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:

    'Assert:

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub




















