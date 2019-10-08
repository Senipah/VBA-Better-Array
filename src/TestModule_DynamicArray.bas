Attribute VB_Name = "TestModule_DynamicArray"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

'@IgnoreModule ProcedureNotUsed
'@IgnoreModule LineLabelNotUsed
'@IgnoreModule EmptyMethod

Private Assert As Object
'@Ignore VariableNotUsed
Private Fakes As Object

Private Const TEST_ARRAY_LENGTH As Long = 10

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
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
    Dim SUT As DynamicArray
    Dim returnedCapacity As Long
    Dim testResult As Boolean
    'Act:
    Set SUT = New DynamicArray
    returnedCapacity = SUT.Capacity
    testResult = (returnedCapacity = 4)
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DynamicArray_Capacity")
Private Sub Capacity_CanSetCapacity_ReturnedCapacityMatchesSetCapacity()
    On Error GoTo TestFail
    
    'Arrange:
    Const DESIRED_CAPACITY As Long = 20
    Dim SUT As DynamicArray
    Dim returnedCapacity As Long
    Dim testResult As Boolean
    Set SUT = New DynamicArray
    
    'Act:
    SUT.Capacity = DESIRED_CAPACITY
    returnedCapacity = SUT.Capacity
    testResult = (returnedCapacity = DESIRED_CAPACITY)
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DynamicArray_Items")
Private Sub Items_CanAssignOneDimemsionalArray_ReturnedArrayEqualsAssignedArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim gen As ArrayGenerator
    Dim SUT As DynamicArray
    Dim testArray As Variant
    Dim returnedItems As Variant
    
    Set gen = New ArrayGenerator
    testArray = gen.getArray(10, VariantVals, OneDimension)
    Set SUT = New DynamicArray
    
    'Act:
    SUT.Items = testArray
    returnedItems = SUT.Items

    'Assert:
    Assert.SequenceEquals testArray, returnedItems

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
    Dim testArray As Variant
    Dim returnedItems As Variant
    
    Set gen = New ArrayGenerator
    testArray = gen.getArray(10, VariantVals, MultiDimension)
    Set SUT = New DynamicArray
    
    'Act:
    SUT.Items = testArray
    returnedItems = SUT.Items

    'Assert:
    Assert.SequenceEquals testArray, returnedItems

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
    Dim testArray As Variant
    Dim i As Long
    Dim j As Long

    Dim returnedItems As Variant
    Dim testResult As Boolean
    
    Set gen = New ArrayGenerator
    testArray = gen.getArray(TEST_ARRAY_LENGTH, VariantVals, Jagged)
    Set SUT = New DynamicArray
    
    'Act:
    SUT.Items = testArray
    returnedItems = SUT.Items
    testResult = True
    For i = LBound(returnedItems) To UBound(returnedItems)
        For j = LBound(returnedItems(i)) To UBound(returnedItems(i))
            If returnedItems(i)(j) <> testArray(i)(j) Then
                testResult = False
                Exit For
            End If
        Next
        If testResult = False Then Exit For
    Next

    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DynamicArray_Length")
Private Sub Length_FromAssignedOneDimensionalArray_ReturnedLengthEqualsOriginalArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim gen As ArrayGenerator
    Dim SUT As DynamicArray
    Dim testArray As Variant
    Dim testResult As Boolean
    
    Set gen = New ArrayGenerator
    testArray = gen.getArray(TEST_ARRAY_LENGTH, VariantVals, OneDimension)
    Set SUT = New DynamicArray
    
    'Act:
    SUT.Items = testArray
    testResult = (SUT.Length = TEST_ARRAY_LENGTH)

    'Assert:
    Assert.IsTrue testResult

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
    Dim testResult As Boolean
    
    Set gen = New ArrayGenerator
    testArray = gen.getArray(TEST_ARRAY_LENGTH, VariantVals, MultiDimension)
    Set SUT = New DynamicArray
    
    'Act:
    SUT.Items = testArray
    testResult = (SUT.Length = TEST_ARRAY_LENGTH)

    'Assert:
    Assert.IsTrue testResult

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
    Dim testResult As Boolean
    
    Set gen = New ArrayGenerator
    testArray = gen.getArray(TEST_ARRAY_LENGTH, VariantVals, Jagged)
    Set SUT = New DynamicArray
    
    'Act:
    SUT.Items = testArray
    testResult = (SUT.Length = TEST_ARRAY_LENGTH)

    'Assert:
    Assert.IsTrue testResult

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
    Dim testResult As Boolean
    
    Set gen = New ArrayGenerator
    testArray = gen.getArray(TEST_ARRAY_LENGTH, VariantVals, OneDimension)
    Set SUT = New DynamicArray
    
    'Act:
    SUT.Items = testArray
    testResult = (SUT.UpperBound = UBound(testArray))

    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Base")
Private Sub Base_FromAssignedOneDimensionalArray_ReturnedBaseEqualsOriginalArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim gen As ArrayGenerator
    Dim SUT As DynamicArray
    Dim testArray As Variant
    Dim testResult As Boolean
    
    Set gen = New ArrayGenerator
    testArray = gen.getArray(TEST_ARRAY_LENGTH, VariantVals, OneDimension)
    Set SUT = New DynamicArray
    
    'Act:
    SUT.Items = testArray
    testResult = (SUT.Base = LBound(testArray))

    'Assert:
    Assert.IsTrue testResult

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
    Dim testResult As Boolean
    
    Dim oldBase As Long
Dim newBase As Long

    
    Set gen = New ArrayGenerator
    testArray = gen.getArray(TEST_ARRAY_LENGTH, VariantVals, OneDimension)
    oldBase = LBound(testArray)
    Set SUT = New DynamicArray
    
    'Act:
    SUT.Items = testArray
    newBase = oldBase + 1
    SUT.Base = newBase
    returnedItems = SUT.Items
    testResult = (LBound(returnedItems) = newBase)

    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Base")
Private Sub Base_ChangingBaseOfAssignedArray_ReturnedArrayHasNewUpperBound()
    On Error GoTo TestFail
    
    'Arrange:
    Dim gen As ArrayGenerator
    Dim SUT As DynamicArray
    Dim testArray As Variant
    Dim returnedItems As Variant
    Dim testResult As Boolean
    
    Dim oldBase As Long
Dim newBase As Long

    
    Set gen = New ArrayGenerator
    testArray = gen.getArray(TEST_ARRAY_LENGTH, VariantVals, OneDimension)
    oldBase = LBound(testArray)
    Set SUT = New DynamicArray
    
    'Act:
    SUT.Items = testArray
    newBase = oldBase + 1
    SUT.Base = newBase
    returnedItems = SUT.Items
    testResult = (UBound(returnedItems) = (UBound(testArray) + 1))

    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Base")
Private Sub Base_ChangingBaseOfAssignedArray_ReturnedArrayHasSameLength()
    On Error GoTo TestFail
    
    'Arrange:
    Dim gen As ArrayGenerator
    Dim SUT As DynamicArray
    Dim testArray As Variant
    Dim testResult As Boolean
    
    Dim oldBase As Long
    Dim expectedBase As Long
    Dim actualBase As Long

    
    Set gen = New ArrayGenerator
    testArray = gen.getArray(TEST_ARRAY_LENGTH, VariantVals, OneDimension)
    oldBase = LBound(testArray)
    Set SUT = New DynamicArray
    
    'Act:
    SUT.Items = testArray
    expectedBase = oldBase + 1
    SUT.Base = expectedBase
    testResult = (SUT.Length = TEST_ARRAY_LENGTH)
    actualBase = SUT.Base
    
    'Assert:
    Assert.IsTrue testResult, "Test result fail"
    Assert.IsTrue (actualBase = expectedBase), "Actual base does not equal expected"
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Item")
Private Sub Item_ChangingExistingIndex_ItemIsChanged()
    On Error GoTo TestFail
    
    'Arrange:
    Const TEST_VALUE As String = "Hello World"
    Dim gen As ArrayGenerator
    Dim SUT As DynamicArray
    Dim testArray As Variant
    Dim testResult As Boolean
    
    Dim actualBase As Long
    Dim expectedBase As Long

    
    Set gen = New ArrayGenerator
    testArray = gen.getArray(TEST_ARRAY_LENGTH, VariantVals, OneDimension)
    expectedBase = LBound(testArray)
    Set SUT = New DynamicArray
    
    'Act:
    SUT.Items = testArray
    SUT.Item(1) = TEST_VALUE
    testResult = (SUT.Item(1) = TEST_VALUE)
    actualBase = SUT.Base

    'Assert:
    Assert.IsTrue testResult, "Test result fail"
    Assert.IsTrue (actualBase = expectedBase), "Actual base does not equal expected base"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Item")
Private Sub Item_ChangingIndexOverUpperBound_ItemIsPushed()
    On Error GoTo TestFail
    
    'Arrange:
    Const TEST_VALUE As String = "Hello World"
    Dim gen As ArrayGenerator
    Dim SUT As DynamicArray
    Dim testArray As Variant
    Dim testResult As Boolean
    
    Dim actualBase As Long
    Dim expectedBase As Long

    
    Set gen = New ArrayGenerator
    testArray = gen.getArray(TEST_ARRAY_LENGTH, VariantVals, OneDimension)
    expectedBase = LBound(testArray)
    Set SUT = New DynamicArray
    
    'Act:
    SUT.Items = testArray
    SUT.Item(SUT.UpperBound + 1) = TEST_VALUE
    testResult = (SUT.Item(SUT.UpperBound) = TEST_VALUE)
    actualBase = SUT.Base
    

    'Assert:
    Assert.IsTrue testResult, "Test result fail"
    Assert.IsTrue (SUT.Length = TEST_ARRAY_LENGTH + 1), "Actual length does not match expected length"
    Assert.IsTrue (actualBase = expectedBase), "Actual base does not match expected base"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Item")
Private Sub Item_ChangingIndexBelowBase_ItemIsUnshifted()
    On Error GoTo TestFail
    
    'Arrange:
    Const TEST_VALUE As String = "Hello World"
    Dim gen As ArrayGenerator
    Dim SUT As DynamicArray
    Dim testArray As Variant
    Dim testResult As Boolean
    
    Dim expectedBase As Long
    Dim actualBase As Long

    
    Set gen = New ArrayGenerator
    testArray = gen.getArray(TEST_ARRAY_LENGTH, VariantVals, OneDimension)
    expectedBase = LBound(testArray)
    Set SUT = New DynamicArray
    
    'Act:
    SUT.Items = testArray
    SUT.Item(SUT.Base - 10) = TEST_VALUE
    testResult = (SUT.Item(SUT.Base) = TEST_VALUE)
    actualBase = SUT.Base

    'Assert:
    Assert.IsTrue testResult, "Actual result does not match expected result"
    Assert.IsTrue (SUT.Length = TEST_ARRAY_LENGTH + 1), "Actual length does not match expected length"
    Assert.IsTrue (actualBase = expectedBase), "Actual base does not match expected base"

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
    Dim returnedItem As Variant
    Dim testResult As Boolean
       
    Set gen = New ArrayGenerator
    testArray = gen.getArray(TEST_ARRAY_LENGTH, VariantVals, OneDimension)
    Set SUT = New DynamicArray
    
    'Act:
    SUT.Items = testArray
    returnedItem = SUT.Item(1)
    testResult = (returnedItem = testArray(1))

    'Assert:
    Assert.IsTrue testResult

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
    Dim returnedObject As Object
       
    Set gen = New ArrayGenerator
    testArray = gen.getArray(TEST_ARRAY_LENGTH, ObjectVals, OneDimension)
    Set SUT = New DynamicArray
    
    'Act:
    SUT.Items = testArray
    Set returnedObject = SUT.Item(1)

    'Assert:
    Assert.AreSame returnedObject, testArray(1)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DynamicArray_Push")
Private Sub Push_AddToNewDynamicArray_ItemAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    Dim Element As Variant
    Element = "Hello World"
    
    'Act:
    SUT.Push Element

    'Assert:
    Assert.IsTrue (SUT.Item(SUT.Base) = Element), "Element value incorrect"
    Assert.IsTrue (SUT.Length = 1), "Length value incorrect"
    Assert.IsTrue (SUT.UpperBound = 0), "Upperbound value incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Push")
Private Sub Push_AddToExistingOneDimensionalArray_ItemAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Dim Element As Variant
    Dim testArray As Variant
    Dim gen As ArrayGenerator
    
    Set SUT = New DynamicArray
    Set gen = New ArrayGenerator
    
    Element = "Hello World"
    testArray = gen.getArray(TEST_ARRAY_LENGTH, VariantVals, OneDimension)
    
    'Act:
    SUT.Items = testArray
    SUT.Push Element

    'Assert:
    Assert.IsTrue (SUT.Item(SUT.UpperBound) = Element), "Element value incorrect"
    Assert.IsTrue (SUT.Length = (TEST_ARRAY_LENGTH + 1)), "Length value incorrect"

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
    Assert.IsTrue (expected = actual), "Element value incorrect"
    Assert.IsTrue (SUT.Length = (TEST_ARRAY_LENGTH + 1)), "Length value incorrect"

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
    Assert.IsTrue (SUT.Length = (TEST_ARRAY_LENGTH + 1)), "Length value incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DynamicArray_Push")
Private Sub Push_AddMultipleToNewDynamicArray_ItemsAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    
    'Act:
    SUT.Push 1, 2, 3

    'Assert:
    Assert.IsTrue (SUT.Item(SUT.Base) = 1), "Element value incorrect"
    Assert.IsTrue (SUT.Length = 3), "Length value incorrect"
    Assert.IsTrue (SUT.UpperBound = 2), "Upperbound value incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


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
    Assert.IsTrue (actual = expected), "Element value incorrect"
    Assert.IsTrue (SUT.Length = (TEST_ARRAY_LENGTH - 1)), "Length value incorrect"
    Assert.IsTrue (SUT.UpperBound = (UBound(testArray) - 1)), "Upperbound value incorrect"
    Assert.IsTrue (actualBase = expectedBase), "Base value incorrect"

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
    Dim actualBase As Long
    Dim expectedBase As Long
    Dim expected As Variant
    Dim actual As Variant
    
    expectedBase = SUT.Base
    expected = Empty
    
    'Act:
    actual = SUT.Pop
    actualBase = SUT.Base

    'Assert:
    Assert.IsTrue (actual = expected), "Element value incorrect"
    Assert.IsTrue (SUT.Length = 0), "Length value incorrect"
    Assert.IsTrue (SUT.UpperBound = 0), "Upperbound value incorrect"
    Assert.IsTrue (actualBase = expectedBase), "Base value incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


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
    Assert.IsTrue (actual = expected), "Element value incorrect"
    Assert.IsTrue (SUT.Length = (TEST_ARRAY_LENGTH - 1)), "Length value incorrect"
    Assert.IsTrue (SUT.UpperBound = (UBound(testArray) - 1)), "Upperbound value incorrect"
    Assert.IsTrue (actualBase = expectedBase), "Base value incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Shift")
Private Sub Shift_ArrayLengthIsZero_ReturnsEmpty()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As DynamicArray
    Set SUT = New DynamicArray
    Dim actualBase As Long
    Dim expectedBase As Long
    Dim expected As Variant
    Dim actual As Variant
    
    expected = Empty
    expectedBase = SUT.Base
    
    'Act:
    actual = SUT.Shift
    actualBase = SUT.Base

    'Assert:
    Assert.IsTrue (actual = expected), "Element value incorrect"
    Assert.IsTrue (SUT.Length = 0), "Length value incorrect"
    Assert.IsTrue (SUT.UpperBound = 0), "Upperbound value incorrect"
    Assert.IsTrue (actualBase = expectedBase), "Base value incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DynamicArray_Unshift")
Private Sub Unshift_OneDimensionalArray_ItemAddedToBeginning()
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
    Assert.IsTrue (actual = expected), "Element value incorrect"
    Assert.IsTrue (SUT.Length = (TEST_ARRAY_LENGTH - 1)), "Length value incorrect"
    Assert.IsTrue (SUT.UpperBound = (UBound(testArray) - 1)), "Upperbound value incorrect"
    Assert.IsTrue (actualBase = expectedBase), "Base value incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DynamicArray_Unshift")
Private Sub Unshift_ArrayLengthIsZero_ItemIsPushedToEmptyArray()

End Sub
