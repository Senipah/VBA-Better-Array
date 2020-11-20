Attribute VB_Name = "TestModule_BetterArray"
Attribute VB_Description = "Unit Tests for 'BetterArray.cls'"
Option Explicit
Option Private Module

'@TestModule
'@Folder("VBABetterArray.Tests")
'@ModuleDescription("Unit Tests for 'BetterArray.cls'")

'@IgnoreModule ProcedureNotUsed
'@IgnoreModule LineLabelNotUsed
'@IgnoreModule EmptyMethod
'@IgnoreModule FunctionReturnValueDiscarded

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

Private Const MISSING_LONG As Long = -9999
Private Const TEST_ARRAY_LENGTH As Long = 10
Private Const EXCEL_DEPENDENCY_WARNING As String = "A test depending on an ExcelProvider instance had failed." & _
        " Once resolved ensure to end any orphan Excel processes running on this system."

' TODO: Ensure test coverage of all paths - ongoing

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
    actual = SUT.lowerBound

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
    SUT.lowerBound = expected
    returnedItems = SUT.Items
    actual = LBound(returnedItems)

    'Assert:
    Assert.AreEqual expected, actual, "Actual LowerBound <> expected"
    Assert.AreEqual SUT.lowerBound, actual, "Actual LowerBound <> SUT.LowerBound prop"
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
    actualLowerBound = SUT.lowerBound

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
    actualLowerBound = SUT.lowerBound
    
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
    SUT.Item(SUT.lowerBound - 10) = expected
    actual = SUT.Item(SUT.lowerBound)
    actualLowerBound = SUT.lowerBound

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

''''''''''''''''''''
' Prop - ArrayType '
''''''''''''''''''''

'Starting Undefined not tested - status not settable externally

' Starting Unalloc - Change to Undefined
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_UnallocToUndefined_ThrowsError()
    Const ExpectedError As Long = ErrorCodes.EC_CANNOT_CONVERT_TO_REQUESTED_STRUCTURE
    On Error GoTo TestFail
    
    'Arrange:
    Dim newType As ArrayTypes
    newType = ArrayTypes.BA_UNDEFINED
    'Act:
    SUT.ArrayType = newType

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

' Starting Unalloc - Change to Unalloc
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_UnallocToUnalloc_Success()
    On Error GoTo TestFail
       
    'Arrange:
    Dim startingType As ArrayTypes
    Dim newType As ArrayTypes
    Dim expected() As Variant
    Dim actual() As Variant
    
    startingType = ArrayTypes.BA_UNALLOCATED
    newType = ArrayTypes.BA_UNALLOCATED
    ReDim expected(SUT.lowerBound)
    
    'Act:
    Assert.AreEqual startingType, SUT.ArrayType
    SUT.ArrayType = newType
    actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Expected undefined to return empty array"
    Assert.AreEqual newType, SUT.ArrayType, "Actual type <> Expected"
    Assert.AreEqual 0&, SUT.Length, "Actual length <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

' Starting Unalloc - Change to OneDimension
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_UnallocToOneDimension_Success()
    On Error GoTo TestFail
       
    'Arrange:
    Dim startingType As ArrayTypes
    Dim newType As ArrayTypes
    Dim expected() As Variant
    Dim actual() As Variant
    
    startingType = ArrayTypes.BA_UNALLOCATED
    newType = ArrayTypes.BA_ONEDIMENSION
    ReDim expected(SUT.lowerBound)
    
    'Act:
    Assert.AreEqual startingType, SUT.ArrayType
    SUT.ArrayType = newType
    actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Expected undefined to return empty array"
    Assert.AreEqual newType, SUT.ArrayType, "Actual type <> Expected"
    Assert.AreEqual 0&, SUT.Length, "Actual length <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


' Starting Unalloc - Change to MultiDimension
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_UnallocToMultiDimension_Success()
    On Error GoTo TestFail
       
    'Arrange:
    Dim startingType As ArrayTypes
    Dim newType As ArrayTypes
    Dim expected() As Variant
    Dim actual() As Variant
    
    startingType = ArrayTypes.BA_UNALLOCATED
    newType = ArrayTypes.BA_MULTIDIMENSION
    ReDim expected(SUT.lowerBound)
    
    'Act:
    Assert.AreEqual startingType, SUT.ArrayType
    SUT.ArrayType = newType
    actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Expected undefined to return empty array"
    Assert.AreEqual newType, SUT.ArrayType, "Actual type <> Expected"
    Assert.AreEqual 0&, SUT.Length, "Actual length <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

' Starting Unalloc - Change to Jagged
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_UnallocToJagged_Success()
    On Error GoTo TestFail
       
    'Arrange:
    Dim startingType As ArrayTypes
    Dim newType As ArrayTypes
    Dim expected() As Variant
    Dim actual() As Variant
    
    startingType = ArrayTypes.BA_UNALLOCATED
    newType = ArrayTypes.BA_JAGGED
    ReDim expected(SUT.lowerBound)
    
    'Act:
    Assert.AreEqual startingType, SUT.ArrayType
    SUT.ArrayType = newType
    actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Expected undefined to return empty array"
    Assert.AreEqual newType, SUT.ArrayType, "Actual type <> Expected"
    Assert.AreEqual 0&, SUT.Length, "Actual length <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


' Starting OneDimension - Change to Undefined
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_OneDimensionToUndefined_ThrowsError()
    Const ExpectedError As Long = ErrorCodes.EC_CANNOT_CONVERT_TO_REQUESTED_STRUCTURE
    On Error GoTo TestFail
    
    'Arrange:
    Dim startingType As ArrayTypes
    Dim newType As ArrayTypes
    Dim testArray() As Variant
    
    startingType = ArrayTypes.BA_ONEDIMENSION
    newType = ArrayTypes.BA_UNDEFINED
    testArray = Gen.GetArray(ArrayType:=startingType)
    SUT.Items = testArray
    
    'Act:
    Assert.AreEqual startingType, SUT.ArrayType
    SUT.ArrayType = newType
    Assert.AreEqual newType, SUT.ArrayType
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

' Starting OneDimension - Change to Unalloc
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_OneDimensionToUnalloc_Success()
    On Error GoTo TestFail
       
    'Arrange:
    Dim startingType As ArrayTypes
    Dim newType As ArrayTypes
    Dim testArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    
    startingType = ArrayTypes.BA_ONEDIMENSION
    newType = ArrayTypes.BA_UNALLOCATED
    ReDim expected(SUT.lowerBound)
    testArray = Gen.GetArray(ArrayType:=startingType)
    SUT.Items = testArray
    
    'Act:
    Assert.AreEqual startingType, SUT.ArrayType
    SUT.ArrayType = newType
    actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Expected undefined to return empty array"
    Assert.AreEqual newType, SUT.ArrayType, "Actual type <> Expected"
    Assert.AreEqual 0&, SUT.Length, "Actual length <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

' Starting OneDimension - Change to OneDimension
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_OneDimensionToOneDimension_Success()
    On Error GoTo TestFail
       
    'Arrange:
    Dim startingType As ArrayTypes
    Dim newType As ArrayTypes
    Dim testArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    
    startingType = ArrayTypes.BA_ONEDIMENSION
    newType = ArrayTypes.BA_ONEDIMENSION
    testArray = Gen.GetArray(ArrayType:=startingType)
    SUT.Items = testArray
    expected = testArray
    
    'Act:
    Assert.AreEqual startingType, SUT.ArrayType
    SUT.ArrayType = newType
    actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Expected undefined to return empty array"
    Assert.AreEqual newType, SUT.ArrayType, "Actual type <> Expected"
    Assert.AreEqual TEST_ARRAY_LENGTH, SUT.Length, "Actual length <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


' Starting OneDimension - Change to MultiDimension
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_OneDimensionToMultiDimension_Success()
    On Error GoTo TestFail
       
    'Arrange:
    Dim startingType As ArrayTypes
    Dim newType As ArrayTypes
    Dim testArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    
    startingType = ArrayTypes.BA_ONEDIMENSION
    newType = ArrayTypes.BA_MULTIDIMENSION
    testArray = Gen.GetArray(ArrayType:=startingType)
    SUT.Items = testArray
    expected = testArray
    
    'Act:
    Assert.AreEqual startingType, SUT.ArrayType
    SUT.ArrayType = newType
    actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Expected undefined to return empty array"
    Assert.AreEqual newType, SUT.ArrayType, "Actual type <> Expected"
    Assert.AreEqual TEST_ARRAY_LENGTH, SUT.Length, "Actual length <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


' Starting OneDimension - Change to Jagged
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_OneDimensionToJagged_Success()
    On Error GoTo TestFail
       
    'Arrange:
    Dim startingType As ArrayTypes
    Dim newType As ArrayTypes
    Dim testArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    
    startingType = ArrayTypes.BA_ONEDIMENSION
    newType = ArrayTypes.BA_JAGGED
    testArray = Gen.GetArray(ArrayType:=startingType)
    SUT.Items = testArray
    expected = testArray
    
    'Act:
    Assert.AreEqual startingType, SUT.ArrayType
    SUT.ArrayType = newType
    actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Expected undefined to return empty array"
    Assert.AreEqual newType, SUT.ArrayType, "Actual type <> Expected"
    Assert.AreEqual TEST_ARRAY_LENGTH, SUT.Length, "Actual length <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

' Starting MultiDimension - Change to Undefined
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_MultiDimensionToUndefined_ThrowsError()
    Const ExpectedError As Long = ErrorCodes.EC_CANNOT_CONVERT_TO_REQUESTED_STRUCTURE
    On Error GoTo TestFail
    
    'Arrange:
    Dim startingType As ArrayTypes
    Dim newType As ArrayTypes
    Dim testArray() As Variant
    
    startingType = ArrayTypes.BA_MULTIDIMENSION
    newType = ArrayTypes.BA_UNDEFINED
    
    testArray = Gen.GetArray(ArrayType:=startingType)
    SUT.Items = testArray
    
    'Act:
    Assert.AreEqual startingType, SUT.ArrayType
    SUT.ArrayType = newType
    Assert.AreEqual newType, SUT.ArrayType

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

' Starting MultiDimension - Change to Unalloc
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_MultiDimensionToUnalloc_Success()
    On Error GoTo TestFail
       
    'Arrange:
    Dim startingType As ArrayTypes
    Dim newType As ArrayTypes
    Dim testArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    
    startingType = ArrayTypes.BA_MULTIDIMENSION
    newType = ArrayTypes.BA_UNALLOCATED
    ReDim expected(SUT.lowerBound)
    testArray = Gen.GetArray(ArrayType:=startingType)
    SUT.Items = testArray
    
    'Act:
    Assert.AreEqual startingType, SUT.ArrayType
    SUT.ArrayType = newType
    actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Expected undefined to return empty array"
    Assert.AreEqual newType, SUT.ArrayType, "Actual type <> Expected"
    Assert.AreEqual 0&, SUT.Length, "Actual length <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

' Starting MultiDimension - Change to OneDimension
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_MultiDimensionToOneDimension_ThrowsError()
    Const ExpectedError As Long = ErrorCodes.EC_CANNOT_CONVERT_TO_REQUESTED_STRUCTURE
    On Error GoTo TestFail
    
    'Arrange:
    Dim startingType As ArrayTypes
    Dim newType As ArrayTypes
    Dim testArray() As Variant
    
    startingType = ArrayTypes.BA_MULTIDIMENSION
    newType = ArrayTypes.BA_ONEDIMENSION
    
    testArray = Gen.GetArray(ArrayType:=startingType)
    SUT.Items = testArray
    
    'Act:
    Assert.AreEqual startingType, SUT.ArrayType
    SUT.ArrayType = newType
    Assert.AreEqual newType, SUT.ArrayType

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

' Starting MultiDimension - Change to MultiDimension
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_MultiDimensionToMultiDimension_Success()
    On Error GoTo TestFail
       
    'Arrange:
    Dim startingType As ArrayTypes
    Dim newType As ArrayTypes
    Dim testArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    
    startingType = ArrayTypes.BA_MULTIDIMENSION
    newType = ArrayTypes.BA_MULTIDIMENSION
    testArray = Gen.GetArray(ArrayType:=startingType)
    SUT.Items = testArray
    expected = testArray
    
    'Act:
    Assert.AreEqual startingType, SUT.ArrayType
    SUT.ArrayType = newType
    actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Expected undefined to return empty array"
    Assert.AreEqual newType, SUT.ArrayType, "Actual type <> Expected"
    Assert.AreEqual TEST_ARRAY_LENGTH, SUT.Length, "Actual length <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


' Starting MultiDimension - Change to Jagged
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_MultiDimensionToJagged_Success()
    On Error GoTo TestFail
       
    'Arrange:
    Dim startingType As ArrayTypes
    Dim newType As ArrayTypes
    Dim testArray() As Variant
    Dim actual() As Variant
    
    startingType = ArrayTypes.BA_MULTIDIMENSION
    newType = ArrayTypes.BA_JAGGED
    testArray = Gen.GetArray(ArrayType:=startingType)
    SUT.Items = testArray
    
    'Act:
    Assert.AreEqual startingType, SUT.ArrayType
    SUT.ArrayType = newType
    actual = SUT.Items
    
    'Assert:
    Assert.IsTrue SequenceEqualsMutiVsJagged(testArray, actual)
    Assert.AreEqual newType, SUT.ArrayType, "Actual type <> Expected"
    Assert.AreEqual TEST_ARRAY_LENGTH, SUT.Length, "Actual length <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


' Starting Jagged - Change to Undefined
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_JaggedToUndefined_ThrowsError()
    Const ExpectedError As Long = ErrorCodes.EC_CANNOT_CONVERT_TO_REQUESTED_STRUCTURE
    On Error GoTo TestFail
    
    'Arrange:
    Dim startingType As ArrayTypes
    Dim newType As ArrayTypes
    Dim testArray() As Variant
    
    startingType = ArrayTypes.BA_JAGGED
    newType = ArrayTypes.BA_UNDEFINED
    
    testArray = Gen.GetArray(ArrayType:=startingType)
    SUT.Items = testArray
    
    'Act:
    Assert.AreEqual startingType, SUT.ArrayType
    SUT.ArrayType = newType
    Assert.AreEqual newType, SUT.ArrayType

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


' Starting Jagged - Change to Unalloc
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_JaggedToUnalloc_Success()
    On Error GoTo TestFail
       
    'Arrange:
    Dim startingType As ArrayTypes
    Dim newType As ArrayTypes
    Dim testArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    
    startingType = ArrayTypes.BA_MULTIDIMENSION
    newType = ArrayTypes.BA_UNALLOCATED
    ReDim expected(SUT.lowerBound)
    testArray = Gen.GetArray(ArrayType:=startingType)
    SUT.Items = testArray
    
    'Act:
    Assert.AreEqual startingType, SUT.ArrayType
    SUT.ArrayType = newType
    actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Expected undefined to return empty array"
    Assert.AreEqual newType, SUT.ArrayType, "Actual type <> Expected"
    Assert.AreEqual 0&, SUT.Length, "Actual length <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

' Starting Jagged - Change to OneDimension
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_JaggedToOneDimension_ThrowsError()
    Const ExpectedError As Long = ErrorCodes.EC_CANNOT_CONVERT_TO_REQUESTED_STRUCTURE
    On Error GoTo TestFail
    
    'Arrange:
    Dim startingType As ArrayTypes
    Dim newType As ArrayTypes
    Dim testArray() As Variant
    
    startingType = ArrayTypes.BA_JAGGED
    newType = ArrayTypes.BA_ONEDIMENSION
    
    testArray = Gen.GetArray(ArrayType:=startingType)
    SUT.Items = testArray
    
    'Act:
    Assert.AreEqual startingType, SUT.ArrayType
    SUT.ArrayType = newType
    Assert.AreEqual newType, SUT.ArrayType
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

' Starting Jagged - Change to MultiDimension
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_JaggedToMultiDimension_Success()
    On Error GoTo TestFail
       
    'Arrange:
    Dim startingType As ArrayTypes
    Dim newType As ArrayTypes
    Dim testArray() As Variant
    Dim actual() As Variant
    
    startingType = ArrayTypes.BA_JAGGED
    newType = ArrayTypes.BA_MULTIDIMENSION
    testArray = Gen.GetArray(ArrayType:=startingType)
    SUT.Items = testArray
    
    'Act:
    Assert.AreEqual startingType, SUT.ArrayType
    SUT.ArrayType = newType
    actual = SUT.Items
    
    'Assert:
    Assert.IsTrue SequenceEqualsMutiVsJagged(actual, testArray)
    Assert.AreEqual newType, SUT.ArrayType, "Actual type <> Expected"
    Assert.AreEqual TEST_ARRAY_LENGTH, SUT.Length, "Actual length <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

' Starting Jagged - Change to Jagged
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_JaggedToJagged_Success()
    On Error GoTo TestFail
       
    'Arrange:
    Dim startingType As ArrayTypes
    Dim newType As ArrayTypes
    Dim testArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    
    startingType = ArrayTypes.BA_JAGGED
    newType = ArrayTypes.BA_JAGGED
    testArray = Gen.GetArray(ArrayType:=startingType)
    SUT.Items = testArray
    expected = testArray
    
    'Act:
    Assert.AreEqual startingType, SUT.ArrayType
    SUT.ArrayType = newType
    actual = SUT.Items
    
    'Assert:
    Assert.IsTrue SequenceEquals_JaggedArray(expected, actual), "Expected undefined to return empty array"
    Assert.AreEqual newType, SUT.ArrayType, "Actual type <> Expected"
    Assert.AreEqual TEST_ARRAY_LENGTH, SUT.Length, "Actual length <> Expected"

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
    actual = SUT.Item(SUT.lowerBound)
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
    actual = SUT.Item(SUT.lowerBound)
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
    expectedLowerBound = SUT.lowerBound
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Pop
    actualLowerBound = SUT.lowerBound

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
    actualLowerBound = SUT.lowerBound
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
    expectedLowerBound = SUT.lowerBound
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Shift
    actualLowerBound = SUT.lowerBound

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
    actualLowerBound = SUT.lowerBound
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
    expectedLowerBound = SUT.lowerBound
    expected = TEST_ARRAY_LENGTH + 1
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Unshift(testElement)
    actualLowerBound = SUT.lowerBound

    'Assert:
    Assert.AreEqual expected, actual, "Return value incorrect"
    Assert.AreEqual (UBound(testArray) + 1), SUT.UpperBound, "Upperbound value incorrect"
    Assert.AreEqual expectedLowerBound, actualLowerBound, "LowerBound value incorrect"
    Assert.AreEqual testElement, SUT.Item(SUT.lowerBound), "Element not inserted at correct position"

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
    actualLowerBound = SUT.lowerBound
    actualUpperBound = SUT.UpperBound
    actualElement = SUT.Item(SUT.lowerBound)

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
    actualLowerBound = SUT.lowerBound
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
    Dim FirstArray() As Variant
    Dim secondArray() As Variant
    
    FirstArray = Gen.GetArray()
    secondArray = Gen.GetArray()
    expected = Gen.ConcatArraysOfSameStructure(AG_ONEDIMENSION, FirstArray, secondArray)
    expectedLength = Gen.GetArrayLength(expected)
    expectedUpperBound = UBound(expected)
    
    'Act:
    SUT.Items = FirstArray
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
    Dim FirstArray() As Variant
    Dim secondArray() As Variant
    Dim actual() As Variant
    Dim expectedLength As Long
    Dim actualLength As Long
    Dim expectedUpperBound As Long
    Dim actualUpperBound As Long
    
    expectedLength = TEST_ARRAY_LENGTH * 2
    FirstArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    secondArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    expected = Gen.ConcatArraysOfSameStructure(AG_MULTIDIMENSION, FirstArray, secondArray)
    expectedUpperBound = UBound(expected)
    
    'Act:
    SUT.Items = FirstArray
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
    Dim FirstArray() As Variant
    Dim secondArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    Dim expectedLength As Long
    Dim actualLength As Long
    Dim expectedUpperBound As Long
    Dim actualUpperBound As Long
    Dim testResult As Boolean
    
    FirstArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    secondArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    
    expected = Gen.ConcatArraysOfSameStructure(AG_JAGGED, FirstArray, secondArray)
    expectedLength = Gen.GetArrayLength(expected)
    expectedUpperBound = UBound(expected)
    
    'Act:
    SUT.Items = FirstArray
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
    Dim FirstArray() As Variant
    Dim secondArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    Dim expectedLength As Long
    Dim actualLength As Long
    Dim expectedUpperBound As Long
    Dim actualUpperBound As Long
    Dim testResult As Boolean
    
    FirstArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    secondArray = Gen.GetArray(ArrayType:=AG_ONEDIMENSION)
    
    expected = Gen.ConcatArraysOfSameStructure(AG_JAGGED, FirstArray, secondArray)
    expectedLength = Gen.GetArrayLength(expected)
    expectedUpperBound = UBound(expected)
    
    'Act:
    SUT.Items = FirstArray
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
    Dim FirstArray() As Variant
    Dim secondArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    Dim expectedLength As Long
    Dim actualLength As Long
    Dim expectedUpperBound As Long
    Dim actualUpperBound As Long
    
    ReDim FirstArray(1 To 2, 1 To 2)
    FirstArray(1, 1) = "Foo"
    FirstArray(1, 2) = "Bar"
    FirstArray(2, 1) = "Fizz"
    FirstArray(2, 2) = "Buzz"
    
    secondArray = Array(1, 2, 3)
    
    ReDim expected(1 To 5, 1 To 2)
    expected(1, 1) = FirstArray(1, 1)
    expected(1, 2) = FirstArray(1, 2)
    expected(2, 1) = FirstArray(2, 1)
    expected(2, 2) = FirstArray(2, 2)
    expected(3, 1) = secondArray(0)
    expected(3, 2) = Empty
    expected(4, 1) = secondArray(1)
    expected(4, 2) = Empty
    expected(5, 1) = secondArray(2)
    expected(5, 2) = Empty
    
    expectedLength = 5
    expectedUpperBound = 5
    
    'Act:
    SUT.Items = FirstArray
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
    Dim FirstArray() As Variant
    Dim secondArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    Dim expectedLength As Long
    Dim actualLength As Long
    Dim expectedUpperBound As Long
    Dim actualUpperBound As Long
    
    FirstArray = Array(1, 2, 3)
    ReDim secondArray(1 To 2, 1 To 2)
    secondArray(1, 1) = "Foo"
    secondArray(1, 2) = "Bar"
    secondArray(2, 1) = "Fizz"
    secondArray(2, 2) = "Buzz"
    
    ReDim expected(0 To 4, 0 To 1)
    expected(0, 0) = FirstArray(0)
    expected(0, 1) = Empty
    expected(1, 0) = FirstArray(1)
    expected(1, 1) = Empty
    expected(2, 0) = FirstArray(2)
    expected(2, 1) = Empty
    expected(3, 0) = secondArray(1, 1)
    expected(3, 1) = secondArray(1, 2)
    expected(4, 0) = secondArray(2, 1)
    expected(4, 1) = secondArray(2, 2)
    
    expectedLength = 5
    expectedUpperBound = 4
    
    'Act:
    SUT.Items = FirstArray
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
Private Sub Concat_AddMultiDimArrayDepth3ToExistingOneDimArray_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim FirstArray() As Variant
    Dim secondArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    Dim expectedLength As Long
    Dim actualLength As Long
    Dim expectedUpperBound As Long
    Dim actualUpperBound As Long
    
    FirstArray = Array(1, 2, 3)
    ReDim secondArray(1 To 2, 1 To 2, 1 To 2)
    secondArray(1, 1, 1) = "Foo"
    secondArray(1, 1, 2) = "Bar"
    secondArray(1, 2, 1) = "Fizz"
    secondArray(1, 2, 2) = "Buzz"
    secondArray(2, 1, 1) = "Foo"
    secondArray(2, 1, 2) = "Bar"
    secondArray(2, 2, 1) = "Fizz"
    secondArray(2, 2, 2) = "Buzz"
    
    ReDim expected(0 To 4, 0 To 1, 0 To 1)
    expected(0, 0, 0) = FirstArray(0)
    expected(1, 0, 0) = FirstArray(1)
    expected(2, 0, 0) = FirstArray(2)
    
    expected(3, 0, 0) = secondArray(1, 1, 1)
    expected(3, 0, 1) = secondArray(1, 1, 2)
    expected(3, 1, 0) = secondArray(1, 2, 1)
    expected(3, 1, 1) = secondArray(1, 2, 2)
    
    expected(4, 0, 0) = secondArray(2, 1, 1)
    expected(4, 0, 1) = secondArray(2, 1, 2)
    expected(4, 1, 0) = secondArray(2, 2, 1)
    expected(4, 1, 1) = secondArray(2, 2, 2)
    
    expectedLength = 5
    expectedUpperBound = 4
    
    'Act:
    SUT.Items = FirstArray
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
    Dim FirstArray() As Variant
    Dim secondArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    Dim expectedLength As Long
    Dim actualLength As Long
    Dim expectedUpperBound As Long
    Dim actualUpperBound As Long
    Dim testResult As Boolean
    
    FirstArray = Gen.GetArray(ArrayType:=AG_ONEDIMENSION)
    secondArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    
    expected = Gen.ConcatArraysOfSameStructure(AG_JAGGED, FirstArray, secondArray)
    expectedLength = Gen.GetArrayLength(expected)
    expectedUpperBound = UBound(expected)
    
    'Act:
    SUT.Items = FirstArray
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
    ReDim expected(SUT.lowerBound)
    actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, (SUT.lowerBound = 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'''''''''''''''''''''''''''''''
' Method - CopyFromCollection '
'''''''''''''''''''''''''''''''

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
    actual = SUT.ToString(prettyPrint:=True)

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
    actual = SUT.ToString(openingDelimiter:="[", closingDelimiter:="]")
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
Private Sub IsSorted_JaggedArrayWithMoreThan2Dimensions_RaisesError()
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

'@TestMethod("BetterArray_Sort")
Private Sub Sort_OneDimArrayObjects_Throws()
    Const ExpectedError As Long = ErrorCodes.EC_CANNOT_SORT_OBJECTS
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    
    testArray = Gen.GetArray(AG_OBJECT)
    SUT.Items = testArray
    'Act:
    SUT.Sort
    
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

'@TestMethod("BetterArray_Sort")
Private Sub Sort_OneDimArrayQuicksortRecursive_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    testArray = Gen.GetArray()
    SUT.Items = testArray
    SUT.SortMethod = SM_QUICKSORT_RECURSIVE
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
Private Sub Sort_OneDimArrayQuicksortRecursiveNegativeBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    testArray = Gen.GetArray()
    SUT.lowerBound = -10
    SUT.Items = testArray
    SUT.SortMethod = SM_QUICKSORT_RECURSIVE
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
Private Sub Sort_OneDimArrayQuicksortRecursivePositiveBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    testArray = Gen.GetArray()
    SUT.lowerBound = 10
    SUT.Items = testArray
    SUT.SortMethod = SM_QUICKSORT_RECURSIVE
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
Private Sub Sort_MultiDimArrayQuicksortRecursive_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    testArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.Items = testArray
    SUT.SortMethod = SM_QUICKSORT_RECURSIVE
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
Private Sub Sort_MultiDimArrayQuicksortRecursiveNegativeBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    testArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.lowerBound = -10
    SUT.Items = testArray
    SUT.SortMethod = SM_QUICKSORT_RECURSIVE
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
Private Sub Sort_MultiDimArrayQuicksortRecursivePositiveBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    testArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.lowerBound = 10
    SUT.Items = testArray
    SUT.SortMethod = SM_QUICKSORT_RECURSIVE
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
Private Sub Sort_JaggedArrayQuicksortRecursive_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    testArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.Items = testArray
    SUT.SortMethod = SM_QUICKSORT_RECURSIVE
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
Private Sub Sort_JaggedArrayQuicksortRecursiveNegativeBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    testArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.lowerBound = -10
    SUT.Items = testArray
    SUT.SortMethod = SM_QUICKSORT_RECURSIVE
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
Private Sub Sort_JaggedArrayQuicksortRecursivePositiveBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    testArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.lowerBound = 10
    SUT.Items = testArray
    SUT.SortMethod = SM_QUICKSORT_RECURSIVE
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
Private Sub Sort_OneDimArrayQuicksortIterative_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    testArray = Gen.GetArray()
    SUT.Items = testArray
    SUT.SortMethod = SM_QUICKSORT_ITERATIVE
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
Private Sub Sort_OneDimArrayQuicksortIterativeNegativeBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    testArray = Gen.GetArray()
    SUT.lowerBound = -10
    SUT.Items = testArray
    SUT.SortMethod = SM_QUICKSORT_ITERATIVE
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
Private Sub Sort_OneDimArrayQuicksortIterativePositiveBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    testArray = Gen.GetArray()
    SUT.lowerBound = 10
    SUT.Items = testArray
    SUT.SortMethod = SM_QUICKSORT_ITERATIVE
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
Private Sub Sort_MultiDimArrayQuicksortIterative_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    testArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.Items = testArray
    SUT.SortMethod = SM_QUICKSORT_ITERATIVE
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
Private Sub Sort_MultiDimArrayQuicksortIterativeNegativeBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    testArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.lowerBound = -10
    SUT.Items = testArray
    SUT.SortMethod = SM_QUICKSORT_ITERATIVE
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
Private Sub Sort_MultiDimArrayQuicksortIterativePositiveBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    testArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.lowerBound = 10
    SUT.Items = testArray
    SUT.SortMethod = SM_QUICKSORT_ITERATIVE
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
Private Sub Sort_JaggedArrayQuicksortIterative_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    testArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.Items = testArray
    SUT.SortMethod = SM_QUICKSORT_ITERATIVE
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
Private Sub Sort_JaggedArrayQuicksortIterativeNegativeBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    testArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.lowerBound = -10
    SUT.Items = testArray
    SUT.SortMethod = SM_QUICKSORT_ITERATIVE
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
Private Sub Sort_JaggedArrayQuicksortIterativePositiveBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    testArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.lowerBound = 10
    SUT.Items = testArray
    SUT.SortMethod = SM_QUICKSORT_ITERATIVE
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
Private Sub Sort_OneDimArrayTimSort_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    testArray = Gen.GetArray()
    SUT.Items = testArray
    SUT.SortMethod = SM_TIMSORT
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
Private Sub Sort_OneDimArrayTimSort10kEntries_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    testArray = Gen.GetArray(Length:=10000)
    SUT.Items = testArray
    SUT.SortMethod = SM_TIMSORT
    'Act:
    SUT.Sort
    actual = SUT.IsSorted
    Dim x() As Variant
    x = SUT.Items
    'Assert:
    Assert.IsTrue actual, "Array not sorted"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_Sort")
Private Sub Sort_OneDimArrayTimSortNegativeBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    testArray = Gen.GetArray()
    SUT.lowerBound = -10
    SUT.Items = testArray
    SUT.SortMethod = SM_TIMSORT
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
Private Sub Sort_OneDimArrayTimSortPositiveBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    testArray = Gen.GetArray()
    SUT.lowerBound = 10
    SUT.Items = testArray
    SUT.SortMethod = SM_TIMSORT
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
Private Sub Sort_MultiDimArrayTimSort_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    testArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.Items = testArray
    SUT.SortMethod = SM_TIMSORT
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
Private Sub Sort_MultiDimArrayTimSortNegativeBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    testArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.lowerBound = -10
    SUT.Items = testArray
    SUT.SortMethod = SM_TIMSORT
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
Private Sub Sort_MultiDimArrayTimSortPositiveBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    testArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.lowerBound = 10
    SUT.Items = testArray
    SUT.SortMethod = SM_TIMSORT
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
Private Sub Sort_JaggedArrayTimSort_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    testArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.Items = testArray
    SUT.SortMethod = SM_TIMSORT
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
Private Sub Sort_JaggedArrayTimSortNegativeBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    testArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.lowerBound = -10
    SUT.Items = testArray
    SUT.SortMethod = SM_TIMSORT
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
Private Sub Sort_JaggedArrayTimSortPositiveBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    testArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.lowerBound = 10
    SUT.Items = testArray
    SUT.SortMethod = SM_TIMSORT
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


'''''''''''''''''''''''
' Method - CopyWithin '
'''''''''''''''''''''''

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
    actual = SUT.Filter("Bar", True, True).Items
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
    SUT.Filter "Bar", False, True
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
    actual = SUT.Filter("Bar", True, True).Items
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'''''''''''''''''''''''
' Method - FilterType '
'''''''''''''''''''''''

'@TestMethod("BetterArray_FilterType")
Private Sub FilterType_OneDimExclude_ReturnsFilteredArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    
    testArray = Array("Foo", 1.23, "Fizz", "Buzz")
    expected = Array("Foo", "Fizz", "Buzz")
    SUT.Items = testArray
    'Act:
    actual = SUT.FilterType("double").Items

    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_FilterType")
Private Sub FilterType_OneDimInclude_ReturnsFilteredArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    
    testArray = Array(1, "Bar", 1.23, 100)
    expected = Array("Bar")
    SUT.Items = testArray
    'Act:
    actual = SUT.FilterType("string", True).Items

    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_FilterType")
Private Sub FilterType_JaggedArrayExclude_ReturnsFilteredArray()
    On Error GoTo TestFail

    'Arrange:
    Dim testArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testResult As Boolean

    testArray = Array(Array("Foo", 1.5), Array("Fizz", "Buzz"))
    expected = Array(Array("Foo"), Array("Fizz", "Buzz"))
    SUT.Items = testArray
    'Act:
    actual = SUT.FilterType("double", False).Items
    testResult = SequenceEquals_JaggedArray(expected, actual)
    'Assert:
    Assert.IsTrue testResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_FilterType")
Private Sub FilterType_JaggedArrayInclude_ReturnsFilteredArray()
    On Error GoTo TestFail

    'Arrange:
    Dim testArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testResult As Boolean

    testArray = Array(Array(1, "Bar"), Array(1.2, -4))
    expected = Array(Array("Bar"))

    SUT.Items = testArray
    'Act:
    actual = SUT.FilterType("string", True, True).Items
    testResult = SequenceEquals_JaggedArray(expected, actual)
    'Assert:
    Assert.IsTrue testResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_FilterType")
Private Sub FilterType_MultiDimArrayExclude_ReturnsFilteredArray()
    On Error GoTo TestFail

    'Arrange:
    Dim testArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant

    ReDim testArray(1 To 2, 1 To 2)
    testArray(1, 1) = "Foo"
    testArray(1, 2) = 1.23
    testArray(2, 1) = "Fizz"
    testArray(2, 2) = "Buzz"

    ReDim expected(1 To 2, 1 To 2)
    expected(1, 1) = "Foo"
    expected(2, 1) = "Fizz"
    expected(2, 2) = "Buzz"

    SUT.Items = testArray
    'Act:
    SUT.FilterType "double", False, True
    actual = SUT.Items

    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_FilterType")
Private Sub FilterType_MultiDimArrayInclude_ReturnsFilteredArray()
    On Error GoTo TestFail

    'Arrange:
    Dim testArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant

    ReDim testArray(1 To 2, 1 To 2)
    testArray(1, 1) = 1.23
    testArray(1, 2) = "Bar"
    testArray(2, 1) = 123
    testArray(2, 2) = 5000

    ReDim expected(1 To 1, 1 To 1)
    expected(1, 1) = "Bar"

    SUT.Items = testArray
    'Act:
    actual = SUT.FilterType("string", True, True).Items
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
    actual = SUT.Includes("Buzz", recurse:=True)
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
    actual = SUT.Includes("wibble", recurse:=True)
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

'''''''''''''''''''''''''
' Method - IncludesType '
'''''''''''''''''''''''''

'@TestMethod("BetterArray_IncludesType")
Private Sub IncludesType_OneDimArrayContainsType_ReturnsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim expected As Boolean
    Dim actual As Boolean
    Dim searchType As String
    testArray = Gen.GetArray(AG_DOUBLE)
    expected = True
    SUT.Items = testArray
    searchType = "Double"
    
    'Act:
    actual = SUT.IncludesType(searchType)
    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_IncludesType")
Private Sub IncludesType_OneDimArrayDoesntContainType_ReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim expected As Boolean
    Dim actual As Boolean
    Dim searchType As String
    testArray = Gen.GetArray(AG_DOUBLE)
    expected = False
    SUT.Items = testArray
    searchType = "string"
    
    'Act:
    actual = SUT.IncludesType(searchType)
    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_IncludesType")
Private Sub IncludesType_JaggedArrayContainsType_ReturnsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim expected As Boolean
    Dim actual As Boolean
    Dim searchType As String
    testArray = Gen.GetArray(AG_DOUBLE)
    expected = True
    SUT.Items = testArray
    searchType = "Double"
    
    'Act:
    actual = SUT.IncludesType(searchType, recurse:=True)
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
    
    SUT.lowerBound = 2
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

'@TestMethod("BetterArray_Max")
Private Sub Max_OneDimArrayNumericInternal_ReturnsLargest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    testArray = Array(1, 3, 2, 6, 4, 9, 0, 5)
    Dim expected As Long
    Dim actual As Long
    
    expected = 9
    SUT.Items = testArray
    'Act:
    actual = SUT.Max

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Max")
Private Sub Max_OneDimArrayStringsInternal_ReturnsLargest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    testArray = Array("Foo", "Bar", "Fizz", "Buzz")
    Dim expected As String
    Dim actual As String
    
    expected = "Foo"
    SUT.Items = testArray
    'Act:
    actual = CStr(SUT.Max)

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Max")
Private Sub Max_OneDimArrayVariantsInternal_ReturnsLargest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    testArray = Array("Foo", 1, "Bar", 100, "Fizz", -1, "Buzz")
    Dim expected As Variant
    Dim actual As Variant
    Dim testResult As Boolean
    
    expected = "Foo"
    SUT.Items = testArray
    'Act:
    actual = SUT.Max
    testResult = ElementsAreEqual(expected, actual)
    'Assert:
    Assert.IsTrue testResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Max")
Private Sub Max_OneDimArrayObjects_ReturnsEmpty()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    testArray = Gen.GetArray(AG_OBJECT)
    Dim expected As Variant
    Dim actual As Variant
    
    expected = Empty
    SUT.Items = testArray
    'Act:
    actual = SUT.Max

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Max")
Private Sub Max_ParamArray_ReturnsLargest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Variant
    Dim actual As Variant
    Dim testResult As Boolean
    
    expected = "Foo"
    'Act:
    actual = SUT.Max("Foo", 1, "Bar", 100, "Fizz", -1, "Buzz")
    testResult = ElementsAreEqual(expected, actual)
    
    'Assert:
    Assert.IsTrue testResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_Max")
Private Sub Max_PassedArray_ReturnsLargest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    testArray = Array("Foo", 1, "Bar", 100, "Fizz", -1, "Buzz")
    Dim expected As Variant
    Dim actual As Variant
    Dim testResult As Boolean
    
    expected = "Foo"
    'Act:
    actual = SUT.Max(testArray)
    testResult = ElementsAreEqual(expected, actual)
    
    'Assert:
    Assert.IsTrue testResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Max")
Private Sub Max_JaggedArray_Returnslargest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Long
    Dim actual As Long
    Dim testArray As Variant
    testArray = Array(Array(1, 3, 20, 4), Array(8, 2, 7, 9))
    expected = 20
    'Act:
    actual = SUT.Max(testArray)

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Max")
Private Sub Max_EmptyInternal_ReturnsEmpty()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Variant
    Dim expected As Variant
    expected = Empty
    
    'Act:
    actual = SUT.Max

    'Assert:
    Assert.AreSame expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


''''''''''''''''
' Method - Min '
''''''''''''''''

'@TestMethod("BetterArray_Min")
Private Sub Min_OneDimArrayNumericInternal_ReturnsSmallest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    testArray = Array(1, 3, 2, 6, 4, 9, 0, 5)
    Dim expected As Long
    Dim actual As Long
    
    expected = 0
    SUT.Items = testArray
    'Act:
    actual = SUT.Min

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Min")
Private Sub Min_OneDimArrayStringsInternal_ReturnsSmallest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    testArray = Array("Foo", "Bar", "Fizz", "Buzz")
    Dim expected As String
    Dim actual As String
    
    expected = "Bar"
    SUT.Items = testArray
    'Act:
    actual = CStr(SUT.Min)

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Min")
Private Sub Min_OneDimArrayVariantsInternal_ReturnsSmallest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    testArray = Array("Foo", 1, "Bar", 100, "Fizz", -1, "Buzz")
    Dim expected As Variant
    Dim actual As Variant
    Dim testResult As Boolean
    
    expected = -1
    SUT.Items = testArray
    'Act:
    actual = SUT.Min
    testResult = ElementsAreEqual(expected, actual)
    'Assert:
    Assert.IsTrue testResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Min")
Private Sub Min_OneDimArrayObjects_ReturnsEmpty()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    testArray = Gen.GetArray(AG_OBJECT)
    Dim expected As Variant
    Dim actual As Variant
    
    expected = Empty
    SUT.Items = testArray
    'Act:
    actual = SUT.Min

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Min")
Private Sub Min_ParamArray_ReturnsSmallest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Variant
    Dim actual As Variant
    Dim testResult As Boolean
    
    expected = -1
    'Act:
    actual = SUT.Min("Foo", 1, "Bar", 100, "Fizz", -1, "Buzz")
    testResult = ElementsAreEqual(expected, actual)
    
    'Assert:
    Assert.IsTrue testResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_Min")
Private Sub Min_PassedArray_ReturnsSmallest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    testArray = Array("Foo", 1, "Bar", 100, "Fizz", -1, "Buzz")
    Dim expected As Variant
    Dim actual As Variant
    Dim testResult As Boolean
    
    expected = -1
    'Act:
    actual = SUT.Min(testArray)
    testResult = ElementsAreEqual(expected, actual)
    
    'Assert:
    Assert.IsTrue testResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Min")
Private Sub Min_JaggedArray_ReturnsSmallest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Long
    Dim actual As Long
    Dim testArray As Variant
    testArray = Array(Array(1, 3, 20, 4), Array(8, 2, 7, 9))
    expected = 1
    'Act:
    actual = SUT.Min(testArray)

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Min")
Private Sub Min_EmptyInternal_ReturnsEmpty()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Variant
    Dim expected As Variant
    expected = Empty
    
    'Act:
    actual = SUT.Min

    'Assert:
    Assert.AreSame expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

''''''''''''''''''
' Method - Slice '
''''''''''''''''''

'@TestMethod("BetterArray_Slice")
Private Sub Slice_OneDimNoEndArg_ReturnsCopy()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    testArray = Gen.GetArray(AG_VARIANT)
    expected = testArray
    
    SUT.Items = testArray
    'Act:
    actual = SUT.Slice(0)
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Slice")
Private Sub Slice_OneDimNoEndArgObjects_ReturnsCopy()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    testArray = Gen.GetArray(AG_OBJECT)
    expected = testArray
    
    SUT.Items = testArray
    'Act:
    actual = SUT.Slice(0)
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Slice")
Private Sub Slice_OneDimWithEndArg_ReturnsCopy()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    testArray = Array("Foo", "Bar", "Fizz", "Buzz")
    expected = Array("Foo", "Bar")
    
    SUT.Items = testArray
    'Act:
    actual = SUT.Slice(0, 2)
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_Slice")
Private Sub Slice_MultiDimNoEndArg_ReturnsCopy()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray(1 To 4, 1 To 2) As Variant
    testArray(1, 1) = "Foo"
    testArray(1, 2) = "Bar"
    testArray(2, 1) = "Fizz"
    testArray(2, 2) = "Buzz"
    testArray(3, 1) = "Xyzzy"
    testArray(3, 2) = "flob"
    testArray(4, 1) = "quux"
    testArray(4, 2) = "quuz"
    
    expected = testArray
    
    SUT.Items = testArray
    'Act:
    actual = SUT.Slice(1)
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Slice")
Private Sub Slice_MultiDimWithEndArg_ReturnsCopy()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected(1 To 2, 1 To 2) As Variant
    Dim actual() As Variant
    Dim testArray(1 To 4, 1 To 2) As Variant
    testArray(1, 1) = "Foo"
    testArray(1, 2) = "Bar"
    testArray(2, 1) = "Fizz"
    testArray(2, 2) = "Buzz"
    testArray(3, 1) = "Xyzzy"
    testArray(3, 2) = "flob"
    testArray(4, 1) = "quux"
    testArray(4, 2) = "quuz"
    
    expected(1, 1) = "Foo"
    expected(1, 2) = "Bar"
    expected(2, 1) = "Fizz"
    expected(2, 2) = "Buzz"
    
    SUT.Items = testArray
    'Act:
    actual = SUT.Slice(1, 3)
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Slice")
Private Sub Slice_JaggedNoEndArg_ReturnsCopy()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testResult As Boolean
    expected = Gen.GetArray(ArrayType:=AG_JAGGED)

    SUT.Items = expected
    'Act:
    actual = SUT.Slice(LBound(expected))
    testResult = SequenceEquals_JaggedArray(expected, actual)
    'Assert:
    Assert.IsTrue testResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Slice")
Private Sub Slice_JaggedWithEndArg_ReturnsCopy()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim testResult As Boolean
    
    testArray = Array(Array("Foo", "Bar"), Array("Fizz", "Buzz"), _
        Array("Xyzzy", "flob"), Array("quux", "quuz"))
   
    expected = Array(Array("Foo", "Bar"), Array("Fizz", "Buzz"))
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Slice(LBound(expected), 2)
    testResult = SequenceEquals_JaggedArray(expected, actual)
    
    'Assert:
    Assert.IsTrue testResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Slice")
Private Sub Slice_EmptyInternal_GracefulDegradation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant

    'Act:
    actual = SUT.Slice(1)
    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

''''''''''''''''''''
' Method - Reverse '
''''''''''''''''''''

'@TestMethod("BetterArray_Reverse")
Private Sub Reverse_OneDimArray_ArrayIsReversed()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testResult As Boolean
    
    expected = Gen.GetArray
    SUT.Items = expected
    
    'Act:
    actual = SUT.Reverse.Items
    testResult = arraysAreReversed(expected, actual)
    'Assert:
    Assert.IsTrue testResult, "Actual not reverse of expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Reverse")
Private Sub Reverse_OneDimArrayBase10_ArrayIsReversed()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testResult As Boolean
    
    Gen.lowerBound = 10
    expected = Gen.GetArray
    SUT.Items = expected
    'Act:
    actual = SUT.Reverse.Items
    testResult = arraysAreReversed(expected, actual)
    'Assert:
    Assert.IsTrue testResult, "Actual not reverse of expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Reverse")
Private Sub Reverse_MultiDimArray_ArrayIsReversed()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testResult As Boolean
    Dim i As Long

    expected = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.Items = expected
    'Act:
    actual = SUT.Reverse.Items
    testResult = True
    For i = LBound(expected) To UBound(expected)
        If Not ElementsAreEqual( _
                expected(i, LBound(expected, 2)), _
                actual(LBound(expected) + UBound(expected) - i, LBound(expected, 2)) _
            ) Then
            testResult = False
            Exit For
        End If
    Next
    
    'Assert:
    Assert.IsTrue testResult, "Actual not reverse of expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Reverse")
Private Sub Reverse_MultiDimArrayRecursive_ArrayIsReversed()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testResult As Boolean
    Dim i As Long
    
    expected = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.Items = expected
    'Act:
    actual = SUT.Reverse(True).Items
    testResult = True
    For i = LBound(expected) To UBound(expected)
        If Not ElementsAreEqual( _
                expected(i, LBound(expected, 2)), _
                actual(LBound(expected) + UBound(expected) - i, UBound(expected, 2)) _
            ) Then
            testResult = False
            Exit For
        End If
    Next
    
    'Assert:
    Assert.IsTrue testResult, "Actual not reverse of expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Reverse")
Private Sub Reverse_JaggedArray_ArrayIsReversed()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testResult As Boolean
    
    expected = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.Items = expected
    'Act:
    actual = SUT.Reverse.Items
    testResult = arraysAreReversed(expected, actual)
    'Assert:
    Assert.IsTrue testResult, "Actual not reverse of expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Reverse")
Private Sub Reverse_JaggedArrayRecurse_ArrayIsReversed()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testResult As Boolean
    
    expected = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.Items = expected
    'Act:
    actual = SUT.Reverse(True).Items
    testResult = arraysAreReversed(expected, actual, True)
    'Assert:
    Assert.IsTrue testResult, "Actual not reverse of expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Reverse")
Private Sub Reverse_EmptyInternal_ReturnsEmpty()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    ReDim expected(0) As Variant
    expected(0) = Empty
    'Act:
    actual = SUT.Reverse.Items

    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub



''''''''''''''''''''
' Method - Shuffle '
''''''''''''''''''''

'@TestMethod("BetterArray_Shuffle")
Private Sub Shuffle_OneDimArray_ArrayIsShuffled()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim sortedArray() As Variant
    Dim actual() As Variant
    
    testArray = Gen.GetArray(AG_DOUBLE)
    SUT.Items = testArray
    sortedArray = SUT.Sort.Items
    'Act:
    actual = SUT.Shuffle.Items

    'Assert:
    Assert.NotSequenceEquals sortedArray, actual, "Array is not shufled"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Shuffle")
Private Sub Shuffle_OneDimArrayBase1_ArrayIsShuffled()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim sortedArray() As Variant
    Dim actual() As Variant
    
    Gen.lowerBound = 1
    testArray = Gen.GetArray(AG_DOUBLE)
    SUT.Items = testArray
    sortedArray = SUT.Sort.Items
    'Act:
    actual = SUT.Shuffle.Items

    'Assert:
    Assert.NotSequenceEquals sortedArray, actual, "Array is not shufled"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Shuffle")
Private Sub Shuffle_MultiDimArray_ArrayIsShuffled()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim sortedArray() As Variant
    Dim actual() As Variant
    
    testArray = Gen.GetArray(AG_DOUBLE, AG_MULTIDIMENSION)
    SUT.Items = testArray
    sortedArray = SUT.Sort.Items
    'Act:
    actual = SUT.Shuffle.Items

    'Assert:
    Assert.NotSequenceEquals sortedArray, actual, "Array is not shufled"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Shuffle")
Private Sub Shuffle_JaggedArray_ArrayIsShuffled()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testArray() As Variant
    Dim sortedArray() As Variant
    Dim actual() As Variant
    Dim testResult As Boolean
    
    testArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.Items = testArray
    sortedArray = SUT.Sort.Items
    'Act:
    actual = SUT.Shuffle.Items
    testResult = SequenceEquals_JaggedArray(sortedArray, actual)
    'Assert:
    Assert.IsFalse testResult, "Array is not shufled"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Shuffle")
Private Sub Shuffle_EmptyInternal_ReturnsEmptyArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    
    'Act:
    ReDim expected(0)
    actual = SUT.Shuffle.Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'''''''''''''''''''''''''''
' Method - FromExcelRange '
'''''''''''''''''''''''''''

'@TestMethod("BetterArray_FromExcelRange")
Private Sub FromExcelRange_NoDetection_Success()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mockData() As Variant
    Gen.lowerBound = 1
    ' Using an array of strings as Excel converts all numbers to double
    mockData = Gen.GetArray(AG_STRING, AG_MULTIDIMENSION)
    
    Dim lastRow As Long
    lastRow = UBound(mockData, 1)
    Dim lastColumn As Long
    lastColumn = UBound(mockData, 2)
    
    Dim ExcelApp As ExcelProvider
    Set ExcelApp = New ExcelProvider
    ExcelApp.CurrentWorksheet.Range("A1").Resize(lastRow, lastColumn).value = mockData
    
    Dim expected(1 To 2, 1 To 2) As Variant
    expected(1, 1) = mockData(1, 1)
    expected(1, 2) = mockData(1, 2)
    expected(2, 1) = mockData(2, 1)
    expected(2, 2) = mockData(2, 2)
    
    Dim actual() As Variant
    
    'Act:
    SUT.FromExcelRange ExcelApp.CurrentWorksheet.Range("A1:B2")
    actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
    
TestExit:
    Exit Sub
TestFail:
    Debug.Print EXCEL_DEPENDENCY_WARNING
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_FromExcelRange")
Private Sub FromExcelRange_ColumnDetection_Success()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mockData() As Variant
    Gen.lowerBound = 1
    ' Using an array of strings as Excel converts all numbers to double
    mockData = Gen.GetArray(AG_STRING, AG_MULTIDIMENSION)
    
    Dim lastRow As Long
    lastRow = UBound(mockData, 1)
    Dim lastColumn As Long
    lastColumn = UBound(mockData, 2)
    
    Dim ExcelApp As ExcelProvider
    Set ExcelApp = New ExcelProvider
    ExcelApp.CurrentWorksheet.Range("A1").Resize(lastRow, lastColumn).value = mockData
    
    Dim i As Long
    Dim expected() As Variant
    ReDim expected(1 To lastRow)
    For i = 1 To lastRow
        expected(i) = mockData(1, i)
    Next
    
    Dim actual() As Variant
    
    'Act:
    SUT.FromExcelRange ExcelApp.CurrentWorksheet.Range("A1"), False, True
    actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
    
TestExit:
    Exit Sub
TestFail:
    Debug.Print EXCEL_DEPENDENCY_WARNING
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_FromExcelRange")
Private Sub FromExcelRange_RowDetection_Success()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mockData() As Variant
    Gen.lowerBound = 1
    ' Using an array of strings as Excel converts all numbers to double
    mockData = Gen.GetArray(AG_STRING, AG_MULTIDIMENSION)
    
    Dim lastRow As Long
    lastRow = UBound(mockData, 1)
    Dim lastColumn As Long
    lastColumn = UBound(mockData, 2)
    
    Dim ExcelApp As ExcelProvider
    Set ExcelApp = New ExcelProvider
    ExcelApp.CurrentWorksheet.Range("A1").Resize(lastRow, lastColumn).value = mockData
    
    Dim i As Long
    Dim expected() As Variant
    ReDim expected(1 To lastRow)
    For i = 1 To lastRow
        expected(i) = mockData(i, 1)
    Next
    
    Dim actual() As Variant
    
    'Act:
    SUT.FromExcelRange ExcelApp.CurrentWorksheet.Range("A1"), True, False
    actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
    
TestExit:
    Exit Sub
TestFail:
    Debug.Print EXCEL_DEPENDENCY_WARNING
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_FromExcelRange")
Private Sub FromExcelRange_ColumnAndRowDetection_Success()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mockData() As Variant
    Gen.lowerBound = 1
    ' Using an array of strings as Excel converts all numbers to double
    mockData = Gen.GetArray(AG_STRING, AG_MULTIDIMENSION)
    
    Dim lastRow As Long
    lastRow = UBound(mockData, 1)
    Dim lastColumn As Long
    lastColumn = UBound(mockData, 2)
    
    Dim ExcelApp As ExcelProvider
    Set ExcelApp = New ExcelProvider
    ExcelApp.CurrentWorksheet.Range("A1").Resize(lastRow, lastColumn).value = mockData
    
    Dim expected() As Variant
    expected = mockData
    
    Dim actual() As Variant
    
    'Act:
    SUT.FromExcelRange ExcelApp.CurrentWorksheet.Range("A1"), True, True
    actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
    
TestExit:
    Exit Sub
TestFail:
    Debug.Print EXCEL_DEPENDENCY_WARNING
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'''''''''''''''''''''''''
' Method - ToExcelRange '
'''''''''''''''''''''''''

'@TestMethod("BetterArray_ToExcelRange")
Private Sub ToExcelRange_OneDimensionNotTransposed_WritesValuesCorrectly()
    On Error GoTo TestFail
    
    'Arrange:
    Dim destination As Object
    Dim returnedRange As Object
    Dim ExcelApp As ExcelProvider
    Dim expected() As Variant
    Dim actual(TEST_ARRAY_LENGTH - 1) As Variant
    
    Set ExcelApp = New ExcelProvider
    Set destination = ExcelApp.CurrentWorksheet.Range("A1")
    
    ' Use Array of Doubles as all values returned from an Excel range are of type Double
    expected = Gen.GetArray(AG_DOUBLE)
    SUT.Items = expected
    
    'Act:
    Set returnedRange = SUT.ToExcelRange(destination)
    Dim i As Long
    For i = 1 To returnedRange.Columns.count
        actual(i - 1) = returnedRange.Cells.Item(1, i).value
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
    Dim destination As Object
    Dim returnedRange As Object
    Dim ExcelApp As ExcelProvider
    Dim expected() As Variant
    Dim actual(TEST_ARRAY_LENGTH - 1) As Variant

    Set ExcelApp = New ExcelProvider
    Set destination = ExcelApp.CurrentWorksheet.Range("A1")
    
    ' Use Array of Doubles as all values returned from an Excel range are of type Double
    expected = Gen.GetArray(AG_DOUBLE)
    SUT.Items = expected
    
    'Act:
    Set returnedRange = SUT.ToExcelRange(destination, True)
    Dim i As Long
    For i = 1 To returnedRange.Rows.count
        actual(i - 1) = returnedRange.Cells.Item(i, 1).value
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
    Dim destination As Object
    Dim ExcelApp As ExcelProvider
    Dim expected() As Variant
    Dim actual As Object
    Dim testResult As Boolean
    Dim transposed As Boolean

    Set ExcelApp = New ExcelProvider
    Set destination = ExcelApp.CurrentWorksheet.Range("A1")
    
    expected = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.Items = expected
    transposed = False
    
    'Act:
    Set actual = SUT.ToExcelRange(destination, transposed)
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
    Dim destination As Object
    Dim ExcelApp As ExcelProvider
    Dim expected() As Variant
    Dim actual As Object
    Dim testResult As Boolean
    Dim transposed As Boolean

    Set ExcelApp = New ExcelProvider
    Set destination = ExcelApp.CurrentWorksheet.Range("A1")
    
    expected = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.Items = expected
    transposed = True
    
    'Act:
    Set actual = SUT.ToExcelRange(destination, transposed)
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
    Dim destination As Object
    Dim returnedRange As Object
    Dim outputSheet As Object
    Dim ExcelApp As ExcelProvider
    Dim i As Long
    Dim j As Long
    Dim expected(TEST_ARRAY_LENGTH - 1, TEST_ARRAY_LENGTH - 1) As Variant
    Dim actual(TEST_ARRAY_LENGTH - 1, TEST_ARRAY_LENGTH - 1) As Variant
    Dim sourceArray() As Variant
    
    Set ExcelApp = New ExcelProvider
    Set outputSheet = ExcelApp.CurrentWorksheet
    Set destination = outputSheet.Range("A1")
    
    ' Use Array of Doubles as all values returned from an Excel range are of type Double
    sourceArray = Gen.GetArray(AG_DOUBLE, AG_JAGGED, Depth:=3)
    
    For i = LBound(sourceArray) To UBound(sourceArray)
        For j = LBound(sourceArray(i)) To UBound(sourceArray(i))
            Set tempBetterArray = New BetterArray
            tempBetterArray.Items = sourceArray(i)(j)
            expected(i, j) = tempBetterArray.ToString()
            Set tempBetterArray = Nothing
        Next
    Next
    
    SUT.Items = sourceArray
    
    'Act:
    Set returnedRange = SUT.ToExcelRange(destination)
    
    For i = 1 To returnedRange.Rows.count
        For j = 1 To returnedRange.Columns.count
            actual(i - 1, j - 1) = returnedRange.Cells.Item(i, j).value
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

'@TestMethod("BetterArray_ParseFromString")
Private Sub ParseFromString_OneDimensionArrayFromToString_ReturnsCorrectValues()
    On Error GoTo TestFail
    
    'Arrange:
    Dim tempBetterArray As BetterArray
    Dim expected() As Variant
    Dim actual() As Variant
    Dim sourceString As String
    Dim testResult As Boolean

    Set tempBetterArray = New BetterArray
    expected = Gen.GetArray()
    tempBetterArray.Items = expected
    sourceString = tempBetterArray.ToString()
    
    'Act:
    actual = SUT.ParseFromString(sourceString).Items
    
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
    Dim sourceString As String
    Dim testResult As Boolean
    
    Set tempBetterArray = New BetterArray
    
    expected = Gen.GetArray(AG_BYTE, AG_JAGGED)
    tempBetterArray.Items = expected
    sourceString = tempBetterArray.ToString()
    
    'Act:
    actual = SUT.ParseFromString(sourceString).Items
    
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
    Dim sourceString As String
    Dim testResult As Boolean
    
    Set tempBetterArray = New BetterArray
    expected = Gen.GetArray(AG_BYTE, AG_JAGGED, Depth:=3)
    tempBetterArray.Items = expected
    sourceString = tempBetterArray.ToString()
    
    'Act:
    actual = SUT.ParseFromString(sourceString).Items
    
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


''''''''''''''''''''''''''''
' Method - Flatten '
''''''''''''''''''''''''''''

'@TestMethod("BetterArray_Flatten")
Private Sub Flatten_OneDimArray_ReturnsSame()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
        
    expected = Gen.GetArray
    SUT.Items = expected
    
    'Act:
    actual = SUT.Flatten.Items
        
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Flatten")
Private Sub Flatten_MultiDimArray_ReturnsFlattenned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected(1 To 4) As Variant
    Dim actual() As Variant
    Dim testArray(1 To 2, 1 To 2) As Variant
    testArray(1, 1) = "Foo"
    testArray(1, 2) = "Bar"
    testArray(2, 1) = "Fizz"
    testArray(2, 2) = "Buzz"
    
    expected(1) = "Foo"
    expected(2) = "Bar"
    expected(3) = "Fizz"
    expected(4) = "Buzz"
    
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Flatten.Items
        
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Flatten")
Private Sub Flatten_JaggedArray_ReturnsFlattenned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected(0 To 3) As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    testArray = Array(Array("Foo", "Bar"), Array("Fizz", "Buzz"))
    
    expected(0) = "Foo"
    expected(1) = "Bar"
    expected(2) = "Fizz"
    expected(3) = "Buzz"
    
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Flatten.Items
        
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_Flatten")
Private Sub Flatten_EmptyInternal_ReturnsArraySizeOne()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected(0) As Variant
    Dim actual() As Variant
    expected(0) = Empty
    
    'Act:
    actual = SUT.Flatten.Items
        
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

''''''''''''''''''
' Method - Clear '
''''''''''''''''''

'@TestMethod("BetterArray_Clear")
Private Sub Clear_OneDimArray_Clears()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected(0) As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim expectedCapacity As Long
    Dim actualCapacity As Long
    
    expected(0) = Empty
    testArray = Gen.GetArray
    SUT.Items = testArray
    expectedCapacity = SUT.Capacity
    'Act:
    actual = SUT.Clear.Items
    actualCapacity = SUT.Capacity
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
    Assert.AreEqual expectedCapacity, actualCapacity, "Actual capacity <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'''''''''''''''''''''''''''
' Method - ResetToDefault '
'''''''''''''''''''''''''''

'@TestMethod("BetterArray_ResetToDefault")
Private Sub ResetToDefault_OneDimArray_Resets()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected(0) As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim expectedCapacity As Long
    Dim actualCapacity As Long
    
    expected(0) = Empty
    testArray = Gen.GetArray
    SUT.Items = testArray
    expectedCapacity = 4
    'Act:
    actual = SUT.ResetToDefault.Items
    actualCapacity = SUT.Capacity
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
    Assert.AreEqual expectedCapacity, actualCapacity, "Actual capacity <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


''''''''''''''''''
' Method - Clone '
''''''''''''''''''

'@TestMethod("BetterArray_Clone")
Private Sub Clone_OneDimArray_CloneIsNotOriginalItemsAreSame()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim clonedSUT As BetterArray
        
    expected = Gen.GetArray
    SUT.Items = expected
    
    'Act:
    Set clonedSUT = SUT.Clone
    actual = clonedSUT.Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
    Assert.AreNotSame SUT, clonedSUT, "Clone is same as original"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'''''''''''''''''''''''''''
' Method - ExtractSegment '
'''''''''''''''''''''''''''

'@TestMethod("BetterArray_ExtractSegment")
Private Sub ExtractSegment_OneDimArrayNoArgs_FullArrayReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
        
    expected = Gen.GetArray
    SUT.Items = expected
    
    'Act:
    actual = SUT.ExtractSegment()
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_ExtractSegment")
Private Sub ExtractSegment_OneDimArrayJustRowArg_ArrayReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim rowIndex As Long
        
    testArray = Gen.GetArray
    SUT.Items = testArray
    rowIndex = 2
    expected = Array(testArray(rowIndex))
    
    'Act:
    actual = SUT.ExtractSegment(rowIndex)
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_ExtractSegment")
Private Sub ExtractSegment_OneDimArrayJustColArg_ArrayReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim columnIndex As Long
        
    testArray = Gen.GetArray
    SUT.Items = testArray
    columnIndex = 3
    expected = Array(testArray(columnIndex))
    
    'Act:
    actual = SUT.ExtractSegment(, columnIndex)
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_ExtractSegment")
Private Sub ExtractSegment_OneDimArrayRowAndColArgs_ArrayReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long
        
    testArray = Gen.GetArray
    SUT.Items = testArray
    rowIndex = 2
    columnIndex = 3
    expected = Array(testArray(rowIndex))
    
    'Act:
    actual = SUT.ExtractSegment(rowIndex, columnIndex)
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_ExtractSegment")
Private Sub ExtractSegment_JaggedArrayNoArgs_FullArrayReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
        
    expected = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.Items = expected
    
    'Act:
    actual = SUT.ExtractSegment()
    
    'Assert:
    Assert.IsTrue SequenceEquals_JaggedArray(expected, actual), "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_ExtractSegment")
Private Sub ExtractSegment_JaggedArrayJustRowArg_ArrayReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim rowIndex As Long
        
    testArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.Items = testArray
    rowIndex = 2
    expected = testArray(rowIndex)
    
    'Act:
    actual = SUT.ExtractSegment(rowIndex)
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_ExtractSegment")
Private Sub ExtractSegment_JaggedArrayJustColArg_ArrayReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim columnIndex As Long
    Dim i As Long
        
    testArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.Items = testArray
    columnIndex = 3
    ReDim expected(LBound(testArray) To UBound(testArray))
    For i = LBound(expected) To UBound(expected)
        expected(i) = testArray(i)(columnIndex)
    Next
    
    'Act:
    actual = SUT.ExtractSegment(, columnIndex)
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_ExtractSegment")
Private Sub ExtractSegment_JaggedDimArrayRowAndColArgs_ArrayReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long
        
    testArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.Items = testArray
    rowIndex = 2
    columnIndex = 3
    expected = Array(testArray(rowIndex)(columnIndex))
    
    'Act:
    actual = SUT.ExtractSegment(rowIndex, columnIndex)
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_ExtractSegment")
Private Sub ExtractSegment_MultiDimArrayNoArgs_FullArrayReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
        
    expected = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.Items = expected
    
    'Act:
    actual = SUT.ExtractSegment()
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_ExtractSegment")
Private Sub ExtractSegment_MultiDimArrayJustRowArg_ArrayReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim rowIndex As Long
    Dim i As Long
        
    testArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.Items = testArray
    rowIndex = 2
    ReDim expected(LBound(testArray, 2) To UBound(testArray, 2))
    For i = LBound(expected) To UBound(expected)
        expected(i) = testArray(rowIndex, i)
    Next
    
    'Act:
    actual = SUT.ExtractSegment(rowIndex)
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_ExtractSegment")
Private Sub ExtractSegment_MultiDimArrayJustColArg_ArrayReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim columnIndex As Long
    Dim i As Long
        
    testArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.Items = testArray
    columnIndex = 3
    ReDim expected(LBound(testArray) To UBound(testArray))
    For i = LBound(expected) To UBound(expected)
        expected(i) = testArray(i, columnIndex)
    Next
    
    'Act:
    actual = SUT.ExtractSegment(, columnIndex)
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_ExtractSegment")
Private Sub ExtractSegment_MultiDimDimArrayRowAndColArgs_ArrayReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long
        
    testArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.Items = testArray
    rowIndex = 2
    columnIndex = 3
    expected = Array(testArray(rowIndex, columnIndex))
    
    'Act:
    actual = SUT.ExtractSegment(rowIndex, columnIndex)
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

''''''''''''''''''''''
' Method - Transpose '
''''''''''''''''''''''
'@TestMethod("BetterArray_Transpose")
Private Sub Transpose_OneDimArray_ArrayTransposed()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim i As Long
        
    testArray = Gen.GetArray()
    SUT.Items = testArray
    
        
    ReDim expected(LBound(testArray) To UBound(testArray), _
        LBound(testArray) To LBound(testArray))
    For i = LBound(testArray) To UBound(testArray)
        expected(i, LBound(testArray)) = testArray(i)
    Next
    
    'Act:
    actual = SUT.Transpose.Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Transpose")
Private Sub Transpose_JaggedArray_ArrayTransposed()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim nested() As Variant
    Dim i As Long
    Dim j As Long
        
    testArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.Items = testArray
    
    ReDim expected(0 To TEST_ARRAY_LENGTH - 1)

    For i = LBound(testArray) To UBound(testArray)
        ReDim nested(0 To TEST_ARRAY_LENGTH - 1)
        For j = LBound(testArray(i)) To UBound(testArray(i))
            nested(j) = testArray(j)(i)
        Next
        expected(i) = nested
    Next
'
    'Act:
    actual = SUT.Transpose.Items
    
    'Assert:
    Assert.IsTrue SequenceEquals_JaggedArray(expected, actual), "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Transpose")
Private Sub Transpose_MultiDimArray_ArrayTransposed()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim i As Long
    Dim j As Long
        
    testArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.Items = testArray
    
    ReDim expected(LBound(testArray, 2) To UBound(testArray, 2), _
        LBound(testArray, 1) To UBound(testArray, 1))
    
    For i = LBound(testArray, 1) To UBound(testArray, 1)
        For j = LBound(testArray, 2) To UBound(testArray, 2)
            expected(j, i) = testArray(i, j)
        Next
    Next
    
    'Act:
    actual = SUT.Transpose.Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

''''''''''''''''''''''
' Method - IndexOf '
''''''''''''''''''''''

'@TestMethod("BetterArray_IndexOf")
Private Sub IndexOf_OneDimArrayValueExists_ReturnsCorrectIndex()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Long
    Dim actual As Long
    Dim testArray() As Variant
    
    expected = 3
        
    testArray = Gen.GetArray()
    SUT.Items = testArray
    
    'Act:
    actual = SUT.IndexOf(testArray(expected))
    
    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_IndexOf")
Private Sub IndexOf_OneDimArrayValueExistsLikeComparison_ReturnsCorrectIndex()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Long
    Dim actual As Long
    Dim testArray() As Variant
    Dim pattern As String
    
    expected = 3
    pattern = "a*a"
    testArray = Array("Zero", "One", "Two", "aBBBa")
    SUT.Items = testArray
    
    'Act:
    actual = SUT.IndexOf(pattern, , CT_LIKENESS)
    
    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_IndexOf")
Private Sub IndexOf_OneDimArrayLikeComparisonPatternNotString_ThrowsError()
    Const ExpectedError As Long = ErrorCodes.EC_STRING_TYPE_EXPECTED
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore VariableNotUsed
    Dim expected As Long
    '@Ignore VariableNotUsed
    Dim actual As Long
    Dim testArray() As Variant
    Dim pattern As Collection
    
    expected = 3
    Set pattern = New Collection
    testArray = Array("Zero", "One", "Two", "aBBBa")
    SUT.Items = testArray
    
    'Act:
    actual = SUT.IndexOf(pattern, , CT_LIKENESS)
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

'@TestMethod("BetterArray_IndexOf")
Private Sub IndexOf_OneDimArrayValueMissing_ReturnsMISSING_LONG()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Long
    Dim actual As Long
    Dim testArray() As Variant
    
    expected = MISSING_LONG
        
    testArray = Gen.GetArray()
    SUT.Items = testArray
    
    'Act:
    actual = SUT.IndexOf("Foo")
    
    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_IndexOf")
Private Sub IndexOf_JaggedArray_ReturnsCorrectIndex()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Long
    Dim actual As Long
    Dim testArray() As Variant
    
    expected = 3
        
    testArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.Items = testArray
    
    'Act:
    actual = SUT.IndexOf(testArray(expected))
    
    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

''''''''''''''''''''''
' Method - Unique '
''''''''''''''''''''''

'@TestMethod("BetterArray_Unique")
Private Sub Unique_OneDimArray_ReturnsUniqueList()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    
    testArray = Array(1, 2, 2, 1, 3, 4, 5, 5, 6, 3)
    expected = Array(1, 2, 3, 4, 5, 6)
        
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Unique.Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_Unique")
Private Sub Unique_JaggedArray_ReturnsUniqueList()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    
    testArray = Array( _
        Array(1, 2, 3), _
        Array("Foo", "Bar"), _
        Array("Foo", "Fizz"), _
        Array(1, 2, 3), _
        Array("Foo", "Bar") _
    )
    expected = Array( _
        Array(1, 2, 3), _
        Array("Foo", "Bar"), _
        Array("Foo", "Fizz") _
    )
        
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Unique.Items
    
    'Assert:
    Assert.IsTrue SequenceEquals_JaggedArray(expected, actual), "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Unique")
Private Sub Unique_JaggedArrayColumnIndexBase0_ReturnsUniqueList()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    
    testArray = Array( _
        Array(1, "Foo", 3), _
        Array(1, "Bar", 3), _
        Array(1, "Foo", 3), _
        Array(1, "Foo", 3), _
        Array(1, "Bar", 3), _
        Array(1, "Fizz", 3), _
        Array(1, "Fizz", 3), _
        Array(1, "Buzz", 3) _
    )
    expected = Array( _
        Array(1, "Foo", 3), _
        Array(1, "Bar", 3), _
        Array(1, "Fizz", 3), _
        Array(1, "Buzz", 3) _
    )
        
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Unique(2).Items
    
    'Assert:
    Assert.IsTrue SequenceEquals_JaggedArray(expected, actual), "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Unique")
Private Sub Unique_JaggedArrayColumnIndexBaseNegativeBase_ReturnsUniqueList()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    
    testArray = Array( _
        Array(1, "Foo", 3), _
        Array(1, "Bar", 3), _
        Array(1, "Foo", 3), _
        Array(1, "Foo", 3), _
        Array(1, "Bar", 3), _
        Array(1, "Fizz", 3), _
        Array(1, "Fizz", 3), _
        Array(1, "Buzz", 3) _
    )
'    Dim expectedRow1(-10 To -7)
'    Dim expectedRow2(-10 To -7)
'    Dim expectedRow3(-10 To -7)
'    expectedRow1
    ReDim expected(-10 To -7)
    expected(-10) = Array(1, "Foo", 3)
    expected(-9) = Array(1, "Bar", 3)
    expected(-8) = Array(1, "Fizz", 3)
    expected(-7) = Array(1, "Buzz", 3)
    
    SUT.lowerBound = -10
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Unique(2).Items
    
    'Assert:
    Assert.IsTrue SequenceEquals_JaggedArray(expected, actual), "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Unique")
Private Sub Unique_JaggedArrayColumnIndexPositiveBase_ReturnsUniqueList()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    
    testArray = Array( _
        Array(1, "Foo", 3), _
        Array(1, "Bar", 3), _
        Array(1, "Foo", 3), _
        Array(1, "Foo", 3), _
        Array(1, "Bar", 3), _
        Array(1, "Fizz", 3), _
        Array(1, "Fizz", 3), _
        Array(1, "Buzz", 3) _
    )
    
    
    ReDim expected(10 To 13)
    expected(10) = Array(1, "Foo", 3)
    expected(11) = Array(1, "Bar", 3)
    expected(12) = Array(1, "Fizz", 3)
    expected(13) = Array(1, "Buzz", 3)
    
    SUT.lowerBound = 10
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Unique(2).Items
    
    'Assert:
    Assert.IsTrue SequenceEquals_JaggedArray(expected, actual), "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'''''''''''''''''''
' Method - Remove '
'''''''''''''''''''

'@TestMethod("BetterArray_Remove")
Private Sub Remove_OneDimArray_RemovesElementAtIndex()
    On Error GoTo TestFail
    
    'Arrange:
    Const removeIndex As Long = 2
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim expectedLength As Long
    Dim actualLength As Long
    
    testArray = Array("Foo", "Bar", "Fizz", "Buzz")
    expected = Array("Foo", "Bar", "Buzz")
    expectedLength = Gen.GetArrayLength(expected)
    
    SUT.Items = testArray
    
    'Act:
    
    actualLength = SUT.Remove(removeIndex)
    actual = SUT.Items
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
    Assert.AreEqual expectedLength, actualLength, "Actual length <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_Remove")
Private Sub Remove_JaggedArray_RemovesElementAtIndex()
    On Error GoTo TestFail
    
    'Arrange:
    Const removeIndex As Long = 2
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim expectedLength As Long
    Dim actualLength As Long
    
    testArray = Array( _
        Array(1, 2, 3), _
        Array("Foo", "Bar"), _
        Array("Foo", "Fizz"), _
        Array(1, 2, 3), _
        Array("Foo", "Bar") _
    )
    expected = Array( _
        Array(1, 2, 3), _
        Array("Foo", "Bar"), _
        Array(1, 2, 3), _
        Array("Foo", "Bar") _
    )
    expectedLength = Gen.GetArrayLength(expected)
    
    SUT.Items = testArray
    
    'Act:
    
    actualLength = SUT.Remove(removeIndex)
    actual = SUT.Items
    'Assert:
    Assert.IsTrue SequenceEquals_JaggedArray(expected, actual), "Actual <> expected"
    Assert.AreEqual expectedLength, actualLength, "Actual length <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_Remove")
Private Sub Remove_MultiDimArray_RemovesElementAtIndex()
    On Error GoTo TestFail
    
    'Arrange:
    Const removeIndex As Long = 2
    Dim expected(1 To 2, 1 To 2) As Variant
    Dim actual() As Variant
    Dim testArray(1 To 3, 1 To 2) As Variant
    Dim expectedLength As Long
    Dim actualLength As Long
    
    testArray(1, 1) = "Foo"
    testArray(1, 2) = "Bar"
    testArray(2, 1) = "Fizz"
    testArray(2, 2) = "Buzz"
    testArray(3, 1) = "Whizz"
    testArray(3, 2) = "Bang"
    
    expected(1, 1) = "Foo"
    expected(1, 2) = "Bar"
    expected(2, 1) = "Whizz"
    expected(2, 2) = "Bang"

    expectedLength = Gen.GetArrayLength(expected)
    
    SUT.Items = testArray
    
    'Act:
    
    actualLength = SUT.Remove(removeIndex)
    actual = SUT.Items
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
    Assert.AreEqual expectedLength, actualLength, "Actual length <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Remove")
Private Sub Remove_OneDimArrayRemoveFirst_RemovesElementAtIndex()
    On Error GoTo TestFail
    
    'Arrange:
    Const removeIndex As Long = 0
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim expectedLength As Long
    Dim actualLength As Long
    
    testArray = Array("Foo", "Bar", "Fizz", "Buzz")
    expected = Array("Bar", "Fizz", "Buzz")
    expectedLength = Gen.GetArrayLength(expected)
    
    SUT.Items = testArray
    
    'Act:
    
    actualLength = SUT.Remove(removeIndex)
    actual = SUT.Items
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
    Assert.AreEqual expectedLength, actualLength, "Actual length <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Remove")
Private Sub Remove_OneDimArrayRemoveLast_RemovesElementAtIndex()
    On Error GoTo TestFail
    
    'Arrange:
    Const removeIndex As Long = 3
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim expectedLength As Long
    Dim actualLength As Long
    
    testArray = Array("Foo", "Bar", "Fizz", "Buzz")
    expected = Array("Foo", "Bar", "Fizz")
    expectedLength = Gen.GetArrayLength(expected)
    
    SUT.Items = testArray
    
    'Act:
    
    actualLength = SUT.Remove(removeIndex)
    actual = SUT.Items
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
    Assert.AreEqual expectedLength, actualLength, "Actual length <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Remove")
Private Sub Remove_OneDimArrayIndexExceedsBounds_NothingRemoved()
    On Error GoTo TestFail
    
    'Arrange:
    Const removeIndex As Long = 100
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim expectedLength As Long
    Dim actualLength As Long
    
    testArray = Array("Foo", "Bar", "Fizz", "Buzz")
    expected = Array("Foo", "Bar", "Fizz", "Buzz")
    expectedLength = Gen.GetArrayLength(expected)
    
    SUT.Items = testArray
    
    'Act:
    
    actualLength = SUT.Remove(removeIndex)
    actual = SUT.Items
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
    Assert.AreEqual expectedLength, actualLength, "Actual length <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

''''''''''''''''''
' Method - Every '
''''''''''''''''''

'@TestMethod("BetterArray_Every")
Private Sub Every_OneDimArrayOfSameString_ReturnsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    testArray = Array("Foo", "Foo", "Foo", "Foo")
    
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Every("Foo")
    
    'Assert:
    Assert.IsTrue actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Every")
Private Sub Every_OneDimArrayOfDifferentString_ReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    testArray = Array("Foo", "Bar", "Foo", "Foo")
    
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Every("Foo")
    
    'Assert:
    Assert.IsFalse actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Every")
Private Sub Every_JaggedDimArrayOfSameString_ReturnsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    testArray = Array(Array("Foo", "Foo", "Foo", "Foo"), Array("Foo", "Foo", "Foo", "Foo"))
    
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Every("Foo")
    
    'Assert:
    Assert.IsTrue actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Every")
Private Sub Every_JaggedDimArrayOfSameString_ReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    testArray = Array(Array("Foo", "Bar", "Foo", "Foo"), Array("Foo", "Foo", "Foo", "Foo"))
    
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Every("Foo")
    
    'Assert:
    Assert.IsFalse actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Every")
Private Sub Every_MiltiDimArrayOfSameString_ReturnsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    ReDim testArray(1 To 2, 1 To 2)
    testArray(1, 1) = "Foo"
    testArray(1, 2) = "Foo"
    testArray(2, 1) = "Foo"
    testArray(2, 2) = "Foo"
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Every("Foo")
    
    'Assert:
    Assert.IsTrue actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Every")
Private Sub Every_MiltiDimArrayOfDifferentString_ReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    ReDim testArray(1 To 2, 1 To 2)
    testArray(1, 1) = "Foo"
    testArray(1, 2) = "Bar"
    testArray(2, 1) = "Foo"
    testArray(2, 2) = "Foo"
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Every("Foo")
    
    'Assert:
    Assert.IsFalse actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'''''''''''''''''''''
' Method - EveryType'
'''''''''''''''''''''

'@TestMethod("BetterArray_EveryType")
Private Sub EveryType_OneDimArrayOfSameString_ReturnsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    testArray = Array("Foo", "Foo", "Foo", "Foo")
    
    SUT.Items = testArray
    
    'Act:
    actual = SUT.EveryType("string")
    
    'Assert:
    Assert.IsTrue actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_EveryType")
Private Sub EveryType_OneDimArrayOfDifferentTypes_ReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    testArray = Array("Foo", 1, 1.2, "Foo")
    
    SUT.Items = testArray
    
    'Act:
    actual = SUT.EveryType("string")
    
    'Assert:
    Assert.IsFalse actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_EveryType")
Private Sub EveryType_JaggedDimArrayOfSameString_ReturnsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    testArray = Array( _
        Array("Foo", "Foo", "Foo", "Foo"), _
        Array("Foo", "Foo", "Foo", "Foo") _
    )
    
    SUT.Items = testArray
    
    'Act:
    actual = SUT.EveryType("string")
    
    'Assert:
    Assert.IsTrue actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_EveryType")
Private Sub EveryType_JaggedDimArrayOfSameString_ReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    testArray = Array( _
        Array("Foo", 1.123, "Foo", "Foo"), _
        Array("Foo", "Foo", "Foo", "Foo") _
    )
    
    SUT.Items = testArray
    
    'Act:
    actual = SUT.EveryType("string")
    
    'Assert:
    Assert.IsFalse actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_EveryType")
Private Sub EveryType_MiltiDimArrayOfSameString_ReturnsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    ReDim testArray(1 To 2, 1 To 2)
    testArray(1, 1) = "Foo"
    testArray(1, 2) = "Foo"
    testArray(2, 1) = "Foo"
    testArray(2, 2) = "Foo"
    SUT.Items = testArray
    
    'Act:
    actual = SUT.EveryType("string")
    
    'Assert:
    Assert.IsTrue actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_EveryType")
Private Sub EveryType_MiltiDimArrayOfDifferentString_ReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Boolean
    Dim testArray() As Variant
    
    ReDim testArray(1 To 2, 1 To 2)
    testArray(1, 1) = "Foo"
    testArray(1, 2) = 1.123
    testArray(2, 1) = "Foo"
    testArray(2, 2) = "Foo"
    SUT.Items = testArray
    
    'Act:
    actual = SUT.EveryType("string")
    
    'Assert:
    Assert.IsFalse actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


''''''''''''''''''''''
' Method - Fill  '
''''''''''''''''''''''

'@TestMethod("BetterArray_Fill")
Private Sub Fill_OneDimArray2To4_SpecifiedIndicesFilled()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual() As Variant
    Dim expected() As Variant
    
    Dim testArray() As Variant
    
    testArray = Gen.GetArray
    
    Const FillVal As Long = 0
        
    expected = testArray
    Dim i As Long
    For i = 2 To 4
        expected(i) = FillVal
    Next
    
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Fill(FillVal, 2, 4).Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_Fill")
Private Sub Fill_OneDimArray1ToEnd_SpecifiedIndicesFilled()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual() As Variant
    Dim expected() As Variant
    
    Dim testArray() As Variant
    
    testArray = Gen.GetArray
    
    Const FillVal As Long = 5
        
    expected = testArray
    Dim i As Long
    For i = 1 To UBound(expected)
        expected(i) = FillVal
    Next
    
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Fill(FillVal, 1).Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_Fill")
Private Sub Fill_OneDimArrayAll_SpecifiedIndicesFilled()
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual() As Variant
    Dim expected() As Variant
    
    Dim testArray() As Variant
    
    testArray = Gen.GetArray
    
    Const FillVal As Long = 6
        
    expected = testArray
    Dim i As Long
    For i = LBound(expected) To UBound(expected)
        expected(i) = FillVal
    Next
    
    SUT.Items = testArray
    
    'Act:
    actual = SUT.Fill(FillVal).Items
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

''''''''''''''''''''''
' Method - LastIndexOf '
''''''''''''''''''''''

'@TestMethod("BetterArray_LastIndexOf")
Private Sub LastIndexOf_OneDimArrayValueExists_ReturnsCorrectIndex()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Long
    Dim actual As Long
    Dim testArray() As Variant
    
    expected = 3
        
    testArray = Array("Dodo", "Tiger", "Penguin", "Dodo")
    SUT.Items = testArray
    
    'Act:
    actual = SUT.LastIndexOf("Dodo")
    
    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_LastIndexOf")
Private Sub LastIndexOf_OneDimArrayValueExistsLikeComparison_ReturnsCorrectIndex()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Long
    Dim actual As Long
    Dim testArray() As Variant
    Dim pattern As String
    
    expected = 3
    pattern = "a*a"
    testArray = Array("Zero", "One", "Two", "aBBBa")
    SUT.Items = testArray
    
    'Act:
    actual = SUT.IndexOf(pattern, , CT_LIKENESS)
    
    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_LastIndexOf")
Private Sub LastIndexOf_OneDimArrayLikeComparisonPatternNotString_ThrowsError()
    Const ExpectedError As Long = ErrorCodes.EC_STRING_TYPE_EXPECTED
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore VariableNotUsed
    Dim expected As Long
    '@Ignore VariableNotUsed
    Dim actual As Long
    Dim testArray() As Variant
    Dim pattern As Collection
    
    expected = 3
    Set pattern = New Collection
    testArray = Array("Zero", "One", "Two", "aBBBa")
    SUT.Items = testArray
    
    'Act:
    actual = SUT.IndexOf(pattern, , CT_LIKENESS)
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

'@TestMethod("BetterArray_LastIndexOf")
Private Sub LastIndexOf_OneDimArrayValueMissing_ReturnsMISSING_LONG()
    On Error GoTo TestFail

    'Arrange:
    Dim expected As Long
    Dim actual As Long
    Dim testArray() As Variant

    expected = MISSING_LONG

    testArray = Gen.GetArray()
    SUT.Items = testArray

    'Act:
    actual = SUT.LastIndexOf("Foo")

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_LastIndexOf")
Private Sub LastIndexOf_JaggedArray_ReturnsCorrectIndex()
    On Error GoTo TestFail

    'Arrange:
    Dim expected As Long
    Dim actual As Long
    Dim testArray() As Variant

    expected = 3

    testArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.Items = testArray

    'Act:
    actual = SUT.LastIndexOf(testArray(expected))

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'''''''''''''''''''
' Method - Splice '
'''''''''''''''''''

'@TestMethod("BetterArray_Splice")
Private Sub Splice_OneDimArrayInsertAtIndex1_Success()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim actualResult() As Variant
    Dim expectedResult() As Variant
    
    expected = Array("Jan", "Feb", "March", "April", "June")
    testArray = Array("Jan", "March", "April", "June")
    SUT.Items = testArray
    ReDim expectedResult(0)

    'Act:
    actualResult = SUT.Splice(1, 0, "Feb")
    actual = SUT.Items

    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
    Assert.SequenceEquals expectedResult, actualResult, "ActualResult <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Splice")
Private Sub Splice_OneDimArrayInsertAtIndex1Delete1_Success()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim actualResult() As Variant
    Dim expectedResult() As Variant
    

    expected = Array("Jan", "Feb", "March", "April", "May")
    testArray = Array("Jan", "Feb", "March", "April", "June")
    SUT.Items = testArray
    expectedResult = Array("June")
    
    'Act:
    actualResult = SUT.Splice(4, 1, "May")
    actual = SUT.Items

    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
    Assert.SequenceEquals expectedResult, actualResult, "ActualResult <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_Splice")
Private Sub Splice_OneDimArrayInsertAtIndex2Delete0Insert2_Success()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim actualResult() As Variant
    Dim expectedResult() As Variant
    

    expected = Array("Banana", "Orange", "Lemon", "Kiwi", "Apple", "Mango")
    testArray = Array("Banana", "Orange", "Apple", "Mango")
    SUT.Items = testArray
    ReDim expectedResult(0)
    
    'Act:
    actualResult = SUT.Splice(2, 0, "Lemon", "Kiwi")
    actual = SUT.Items

    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
    Assert.SequenceEquals expectedResult, actualResult, "ActualResult <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_Splice")
Private Sub Splice_OneDimArrayInsertAtIndex2Delete1Insert2_Success()
    On Error GoTo TestFail
        
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim actualResult() As Variant
    Dim expectedResult() As Variant
    

    expected = Array("Banana", "Orange", "Lemon", "Kiwi", "Mango")
    testArray = Array("Banana", "Orange", "Apple", "Mango")
    SUT.Items = testArray
    expectedResult = Array("Apple")
    
    'Act:
    actualResult = SUT.Splice(2, 1, "Lemon", "Kiwi")
    actual = SUT.Items

    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
    Assert.SequenceEquals expectedResult, actualResult, "ActualResult <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_Splice")
Private Sub Splice_OneDimArrayInsertAtIndex2Delete2Insert0_Success()
    On Error GoTo TestFail
        
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim actualResult() As Variant
    Dim expectedResult() As Variant
    

    expected = Array("Banana", "Orange", "Kiwi")
    testArray = Array("Banana", "Orange", "Apple", "Mango", "Kiwi")
    SUT.Items = testArray
    expectedResult = Array("Apple", "Mango")
    
    'Act:
    actualResult = SUT.Splice(2, 2)
    actual = SUT.Items

    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
    Assert.SequenceEquals expectedResult, actualResult, "ActualResult <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_Splice")
Private Sub Splice_OneDimArrayBase1InsertAtIndex2Delete0Insert2_Success()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim actualResult() As Variant
    Dim expectedResult() As Variant
        
    SUT.lowerBound = 1
    
    ReDim expected(1 To 6)
    expected(1) = "Banana"
    expected(2) = "Orange"
    expected(3) = "Lemon"
    expected(4) = "Kiwi"
    expected(5) = "Apple"
    expected(6) = "Mango"
    
    testArray = Array("Banana", "Orange", "Apple", "Mango")
    SUT.Items = testArray
    ReDim expectedResult(0)
    
    'Act:
    actualResult = SUT.Splice(3, 0, "Lemon", "Kiwi")
    actual = SUT.Items

    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
    Assert.SequenceEquals expectedResult, actualResult, "ActualResult <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_Splice")
Private Sub Splice_OneDimArrayBase1InsertAtIndex2Delete1Insert2_Success()
    On Error GoTo TestFail
        
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim actualResult() As Variant
    Dim expectedResult() As Variant
    
    SUT.lowerBound = 1
    ReDim expected(1 To 5)
    expected(1) = "Banana"
    expected(2) = "Orange"
    expected(3) = "Lemon"
    expected(4) = "Kiwi"
    expected(5) = "Mango"
    
    testArray = Array("Banana", "Orange", "Apple", "Mango")
    SUT.Items = testArray
    expectedResult = Array("Apple")
    
    'Act:
    actualResult = SUT.Splice(3, 1, "Lemon", "Kiwi")
    actual = SUT.Items

    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
    Assert.SequenceEquals expectedResult, actualResult, "ActualResult <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_Splice")
Private Sub Splice_OneDimArrayBase1InsertAtIndex2Delete2Insert0_Success()
    On Error GoTo TestFail
        
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim actualResult() As Variant
    Dim expectedResult() As Variant
    
    SUT.lowerBound = 1
    
    ReDim expected(1 To 3)
    expected(1) = "Banana"
    expected(2) = "Orange"
    expected(3) = "Kiwi"
    
    testArray = Array("Banana", "Orange", "Apple", "Mango", "Kiwi")
    SUT.Items = testArray
    expectedResult = Array("Apple", "Mango")
    
    'Act:
    actualResult = SUT.Splice(3, 2)
    actual = SUT.Items

    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
    Assert.SequenceEquals expectedResult, actualResult, "ActualResult <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_Splice")
Private Sub Splice_OneDimArrayBaseNegative1InsertAtIndex2Delete0Insert2_Success()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim actualResult() As Variant
    Dim expectedResult() As Variant
    
    SUT.lowerBound = -1
    
    ReDim expected(-1 To 4)
    expected(-1) = "Banana"
    expected(0) = "Orange"
    expected(1) = "Lemon"
    expected(2) = "Kiwi"
    expected(3) = "Apple"
    expected(4) = "Mango"

    testArray = Array("Banana", "Orange", "Apple", "Mango")
    SUT.Items = testArray
    ReDim expectedResult(0)
    
    'Act:
    actualResult = SUT.Splice(1, 0, "Lemon", "Kiwi")
    actual = SUT.Items

    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
    Assert.SequenceEquals expectedResult, actualResult, "ActualResult <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_Splice")
Private Sub Splice_OneDimArrayBaseNegative1InsertAtIndex2Delete1Insert2_Success()
    On Error GoTo TestFail
        
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim actualResult() As Variant
    Dim expectedResult() As Variant
    
    SUT.lowerBound = -1
    ReDim expected(-1 To 3)
    expected(-1) = "Banana"
    expected(0) = "Orange"
    expected(1) = "Lemon"
    expected(2) = "Kiwi"
    expected(3) = "Mango"
    
    testArray = Array("Banana", "Orange", "Apple", "Mango")
    SUT.Items = testArray
    expectedResult = Array("Apple")
    
    'Act:
    actualResult = SUT.Splice(1, 1, "Lemon", "Kiwi")
    actual = SUT.Items

    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
    Assert.SequenceEquals expectedResult, actualResult, "ActualResult <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_Splice")
Private Sub Splice_OneDimArrayBaseNegative1InsertAtIndex2Delete2Insert0_Success()
    On Error GoTo TestFail
        
    'Arrange:
    Dim expected() As Variant
    Dim actual() As Variant
    Dim testArray() As Variant
    Dim actualResult() As Variant
    Dim expectedResult() As Variant
    
    SUT.lowerBound = -1
    ReDim expected(-1 To 1)
    expected(-1) = "Banana"
    expected(0) = "Orange"
    expected(1) = "Kiwi"
    
    testArray = Array("Banana", "Orange", "Apple", "Mango", "Kiwi")
    SUT.Items = testArray
    expectedResult = Array("Apple", "Mango")
    
    'Act:
    actualResult = SUT.Splice(1, 2)
    actual = SUT.Items

    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> expected"
    Assert.SequenceEquals expectedResult, actualResult, "ActualResult <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


''''''''''''''''''''''''''
' Method - FromCSVString '
''''''''''''''''''''''''''

'@TestMethod("BetterArray_FromCSVString")
Private Sub FromCSVString_Simple10RowWithHeaders_ReturnsJagged()
    On Error GoTo TestFail
        
    'Arrange:
    Const TEST_DATA As String = _
        "Region,Country A,Item Type,Sales Channel,Order Priority,Order Date,Order ID,Ship Date,Units Sold,Unit Price,Unit Cost,Total Revenue,Total Cost,Total Profit" & vbCrLf & _
        "Sub-Saharan Africa,Chad,Office Supplies,Online,L,1/27/2011,292494523,2/12/2011,4484,651.21,524.96,2920025.64,2353920.64,566105.00" & vbCrLf & _
        "Europe , Latvia, Beverages, Online, C, 12 / 28 / 2015, 361825549, 1 / 23 / 2016, 1075, 47.45, 31.79, 51008.75, 34174.25, 16834.5" & vbCrLf & _
        "Middle East and North Africa,Pakistan,Vegetables,Offline,C,1/13/2011,141515767,2/1/2011,6515,154.06,90.93,1003700.90,592408.95,411291.95" & vbCrLf & _
        "Sub-Saharan Africa,Democratic Republic of the Congo,Household,Online,C,9/11/2012,500364005,10/6/2012,7683,668.27,502.54,5134318.41,3861014.82,1273303.59" & vbCrLf & _
        "Europe,Czech Republic,Beverages,Online,C,10/27/2015,127481591,12/5/2015,3491,47.45,31.79,165647.95,110978.89,54669.06" & vbCrLf & _
        "Sub-Saharan Africa,South Africa,Beverages,Offline,H,7/10/2012,482292354,8/21/2012,9880,47.45,31.79,468806.00,314085.20,154720.80" & vbCrLf & _
        "Asia , Laos, Vegetables, Online, L, 2 / 20 / 2011, 844532620, 3 / 20 / 2011, 4825, 154.06, 90.93, 743339.50, 438737.25, 304602.25" & vbCrLf & _
        "Asia,China,Baby Food,Online,C,4/10/2017,564251220,5/12/2017,3330,255.28,159.42,850082.40,530868.60,319213.80" & vbCrLf & _
        "Sub-Saharan Africa,Eritrea,Meat,Online,L,11/21/2014,411809480,1/10/2015,2431,421.89,364.69,1025614.59,886561.39,139053.20"

    Dim expected() As Variant
    Dim actual() As Variant
    ReDim expected(0 To 9)

    expected(0) = Array("Region", "Country A", "Item Type", "Sales Channel", "Order Priority", "Order Date", "Order ID", "Ship Date", "Units Sold", "Unit Price", "Unit Cost", "Total Revenue", "Total Cost", "Total Profit")
    expected(1) = Array("Sub-Saharan Africa", "Chad", "Office Supplies", "Online", "L", "1/27/2011", "292494523", "2/12/2011", "4484", "651.21", "524.96", "2920025.64", "2353920.64", "566105.00")
    expected(2) = Array("Europe", "Latvia", "Beverages", "Online", "C", "12 / 28 / 2015", "361825549", "1 / 23 / 2016", "1075", "47.45", "31.79", "51008.75", "34174.25", "16834.5")
    expected(3) = Array("Middle East and North Africa", "Pakistan", "Vegetables", "Offline", "C", "1/13/2011", "141515767", "2/1/2011", "6515", "154.06", "90.93", "1003700.90", "592408.95", "411291.95")
    expected(4) = Array("Sub-Saharan Africa", "Democratic Republic of the Congo", "Household", "Online", "C", "9/11/2012", "500364005", "10/6/2012", "7683", "668.27", "502.54", "5134318.41", "3861014.82", "1273303.59")
    expected(5) = Array("Europe", "Czech Republic", "Beverages", "Online", "C", "10/27/2015", "127481591", "12/5/2015", "3491", "47.45", "31.79", "165647.95", "110978.89", "54669.06")
    expected(6) = Array("Sub-Saharan Africa", "South Africa", "Beverages", "Offline", "H", "7/10/2012", "482292354", "8/21/2012", "9880", "47.45", "31.79", "468806.00", "314085.20", "154720.80")
    expected(7) = Array("Asia", "Laos", "Vegetables", "Online", "L", "2 / 20 / 2011", "844532620", "3 / 20 / 2011", "4825", "154.06", "90.93", "743339.50", "438737.25", "304602.25")
    expected(8) = Array("Asia", "China", "Baby Food", "Online", "C", "4/10/2017", "564251220", "5/12/2017", "3330", "255.28", "159.42", "850082.40", "530868.60", "319213.80")
    expected(9) = Array("Sub-Saharan Africa", "Eritrea", "Meat", "Online", "L", "11/21/2014", "411809480", "1/10/2015", "2431", "421.89", "364.69", "1025614.59", "886561.39", "139053.20")
           
    'Act:
    actual = SUT.FromCSVString(TEST_DATA).Items

    'Assert:
    Assert.IsTrue SequenceEquals_JaggedArray(expected, actual), "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_FromCSVString")
Private Sub FromCSVString_RFC4180_ReturnsJagged()
    On Error GoTo TestFail
        
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim line1 As String
    '@Ignore UseMeaningfulName
    Dim line2 As String
    '@Ignore UseMeaningfulName
    Dim line3 As String
    Dim CSVData As String
    line1 = _
        WrapQuote("Field with " & vbCrLf & "multiple lines") & " ," & _
        WrapQuote("Another field " & vbCrLf & "with some " & vbCrLf & "line breaks inside") & " , " & _
        WrapQuote("Include some  comma, for test, and some [" & WrapQuote() & "] Quotes") & " , " & _
        WrapQuote("Normal field here") & vbCrLf
    line2 = "1, 2, 3 ,4 " & vbCrLf
    line3 = "Field 1, Field 2 , Field 3 , Field 4"
    CSVData = line1 & line2 & line3

    Dim expected() As Variant
    Dim actual() As Variant
    ReDim expected(0 To 2)

    expected(0) = Array( _
        "Field with " & vbCrLf & "multiple lines", _
        "Another field " & vbCrLf & "with some " & vbCrLf & "line breaks inside", _
        "Include some  comma, for test, and some [" & WrapQuote() & "] Quotes", _
        "Normal field here" _
    )
    expected(1) = Array("1", "2", "3", "4")
    expected(2) = Array("Field 1", "Field 2", "Field 3", "Field 4")
           
    'Act:
    actual = SUT.FromCSVString(CSVData).Items

    'Assert:
    Assert.IsTrue SequenceEquals_JaggedArray(expected, actual), "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BetterArray_FromCSVString")
Private Sub FromCSVString_NullString_ReturnsJagged()
    On Error GoTo TestFail
        
    'Arrange:
    Dim CSVData As String
    CSVData = _
        WrapQuote() & "," & vbCrLf & _
        "," & WrapQuote() & " " & vbCrLf & _
        "Field1,Field2" & vbCrLf

    Dim expected() As Variant
    Dim actual() As Variant
    ReDim expected(0 To 2)

    expected(0) = Array(vbNullString, vbNullString)
    expected(1) = Array(vbNullString, vbNullString)
    expected(2) = Array("Field1", "Field2")
           
    'Act:
    actual = SUT.FromCSVString(CSVData).Items

    'Assert:
    Assert.IsTrue SequenceEquals_JaggedArray(expected, actual), "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'''''''''''''''
' ToCSVString '
'''''''''''''''

'@TestMethod("BetterArray_ToCSVString")
Private Sub ToCSVString_Simple10RowWithHeaders_ValidStringReturned()
    On Error GoTo TestFail

    'Arrange:
    Const expected As String = _
        "Region,Country,Item Type,Sales Channel,Order Priority,Order Date,Order ID,Ship Date,Units Sold,Unit Price,Unit Cost,Total Revenue,Total Cost,Total Profit" & vbCrLf & _
        "Sub-Saharan Africa,Chad,Office Supplies,Online,L,1/27/2011,292494523,2/12/2011,4484,651.21,524.96,2920025.64,2353920.64,566105.00" & vbCrLf & _
        "Europe,Latvia,Beverages,Online,C,12/28/2015,361825549,1/23/2016,1075,47.45,31.79,51008.75,34174.25,16834.50" & vbCrLf & _
        "Middle East and North Africa,Pakistan,Vegetables,Offline,C,1/13/2011,141515767,2/1/2011,6515,154.06,90.93,1003700.90,592408.95,411291.95" & vbCrLf & _
        "Sub-Saharan Africa,Democratic Republic of the Congo,Household,Online,C,9/11/2012,500364005,10/6/2012,7683,668.27,502.54,5134318.41,3861014.82,1273303.59" & vbCrLf & _
        "Europe,Czech Republic,Beverages,Online,C,10/27/2015,127481591,12/5/2015,3491,47.45,31.79,165647.95,110978.89,54669.06" & vbCrLf & _
        "Sub-Saharan Africa,South Africa,Beverages,Offline,H,7/10/2012,482292354,8/21/2012,9880,47.45,31.79,468806.00,314085.20,154720.80" & vbCrLf & _
        "Asia,Laos,Vegetables,Online,L,2/20/2011,844532620,3/20/2011,4825,154.06,90.93,743339.50,438737.25,304602.25" & vbCrLf & _
        "Asia,China,Baby Food,Online,C,4/10/2017,564251220,5/12/2017,3330,255.28,159.42,850082.40,530868.60,319213.80" & vbCrLf & _
        "Sub-Saharan Africa,Eritrea,Meat,Online,L,11/21/2014,411809480,1/10/2015,2431,421.89,364.69,1025614.59,886561.39,139053.20"

    Dim headers() As Variant
    Dim testDatum() As Variant
    Dim actual As String
    ReDim testDatum(0 To 8)

    headers = Array("Region", "Country", "Item Type", "Sales Channel", "Order Priority", "Order Date", "Order ID", "Ship Date", "Units Sold", "Unit Price", "Unit Cost", "Total Revenue", "Total Cost", "Total Profit")
    testDatum(0) = Array("Sub-Saharan Africa", "Chad", "Office Supplies", "Online", "L", "1/27/2011", "292494523", "2/12/2011", "4484", "651.21", "524.96", "2920025.64", "2353920.64", "566105.00")
    testDatum(1) = Array("Europe", "Latvia", "Beverages", "Online", "C", "12/28/2015", "361825549", "1/23/2016", "1075", "47.45", "31.79", "51008.75", "34174.25", "16834.50")
    testDatum(2) = Array("Middle East and North Africa", "Pakistan", "Vegetables", "Offline", "C", "1/13/2011", "141515767", "2/1/2011", "6515", "154.06", "90.93", "1003700.90", "592408.95", "411291.95")
    testDatum(3) = Array("Sub-Saharan Africa", "Democratic Republic of the Congo", "Household", "Online", "C", "9/11/2012", "500364005", "10/6/2012", "7683", "668.27", "502.54", "5134318.41", "3861014.82", "1273303.59")
    testDatum(4) = Array("Europe", "Czech Republic", "Beverages", "Online", "C", "10/27/2015", "127481591", "12/5/2015", "3491", "47.45", "31.79", "165647.95", "110978.89", "54669.06")
    testDatum(5) = Array("Sub-Saharan Africa", "South Africa", "Beverages", "Offline", "H", "7/10/2012", "482292354", "8/21/2012", "9880", "47.45", "31.79", "468806.00", "314085.20", "154720.80")
    testDatum(6) = Array("Asia", "Laos", "Vegetables", "Online", "L", "2/20/2011", "844532620", "3/20/2011", "4825", "154.06", "90.93", "743339.50", "438737.25", "304602.25")
    testDatum(7) = Array("Asia", "China", "Baby Food", "Online", "C", "4/10/2017", "564251220", "5/12/2017", "3330", "255.28", "159.42", "850082.40", "530868.60", "319213.80")
    testDatum(8) = Array("Sub-Saharan Africa", "Eritrea", "Meat", "Online", "L", "11/21/2014", "411809480", "1/10/2015", "2431", "421.89", "364.69", "1025614.59", "886561.39", "139053.20")

    'Act:
    SUT.Items = testDatum
    actual = SUT.ToCSVString(headers:=headers)
    
    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_ToCSVString")
Private Sub ToCSVString_Simple10RowNoHeaders_ValidStringReturned()
    On Error GoTo TestFail

    'Arrange:
    Const expected As String = _
        "Region,Country,Item Type,Sales Channel,Order Priority,Order Date,Order ID,Ship Date,Units Sold,Unit Price,Unit Cost,Total Revenue,Total Cost,Total Profit" & vbCrLf & _
        "Sub-Saharan Africa,Chad,Office Supplies,Online,L,1/27/2011,292494523,2/12/2011,4484,651.21,524.96,2920025.64,2353920.64,566105.00" & vbCrLf & _
        "Europe,Latvia,Beverages,Online,C,12/28/2015,361825549,1/23/2016,1075,47.45,31.79,51008.75,34174.25,16834.50" & vbCrLf & _
        "Middle East and North Africa,Pakistan,Vegetables,Offline,C,1/13/2011,141515767,2/1/2011,6515,154.06,90.93,1003700.90,592408.95,411291.95" & vbCrLf & _
        "Sub-Saharan Africa,Democratic Republic of the Congo,Household,Online,C,9/11/2012,500364005,10/6/2012,7683,668.27,502.54,5134318.41,3861014.82,1273303.59" & vbCrLf & _
        "Europe,Czech Republic,Beverages,Online,C,10/27/2015,127481591,12/5/2015,3491,47.45,31.79,165647.95,110978.89,54669.06" & vbCrLf & _
        "Sub-Saharan Africa,South Africa,Beverages,Offline,H,7/10/2012,482292354,8/21/2012,9880,47.45,31.79,468806.00,314085.20,154720.80" & vbCrLf & _
        "Asia,Laos,Vegetables,Online,L,2/20/2011,844532620,3/20/2011,4825,154.06,90.93,743339.50,438737.25,304602.25" & vbCrLf & _
        "Asia,China,Baby Food,Online,C,4/10/2017,564251220,5/12/2017,3330,255.28,159.42,850082.40,530868.60,319213.80" & vbCrLf & _
        "Sub-Saharan Africa,Eritrea,Meat,Online,L,11/21/2014,411809480,1/10/2015,2431,421.89,364.69,1025614.59,886561.39,139053.20"
        
        
    Dim testDatum() As Variant
    Dim actual As String
    ReDim testDatum(0 To 9)

    testDatum(0) = Array("Region", "Country", "Item Type", "Sales Channel", "Order Priority", "Order Date", "Order ID", "Ship Date", "Units Sold", "Unit Price", "Unit Cost", "Total Revenue", "Total Cost", "Total Profit")
    testDatum(1) = Array("Sub-Saharan Africa", "Chad", "Office Supplies", "Online", "L", "1/27/2011", "292494523", "2/12/2011", "4484", "651.21", "524.96", "2920025.64", "2353920.64", "566105.00")
    testDatum(2) = Array("Europe", "Latvia", "Beverages", "Online", "C", "12/28/2015", "361825549", "1/23/2016", "1075", "47.45", "31.79", 51008.75, "34174.25", "16834.50")
    testDatum(3) = Array("Middle East and North Africa", "Pakistan", "Vegetables", "Offline", "C", "1/13/2011", "141515767", "2/1/2011", "6515", "154.06", "90.93", "1003700.90", "592408.95", "411291.95")
    testDatum(4) = Array("Sub-Saharan Africa", "Democratic Republic of the Congo", "Household", "Online", "C", "9/11/2012", "500364005", "10/6/2012", "7683", "668.27", "502.54", "5134318.41", "3861014.82", "1273303.59")
    testDatum(5) = Array("Europe", "Czech Republic", "Beverages", "Online", "C", "10/27/2015", "127481591", "12/5/2015", "3491", "47.45", "31.79", "165647.95", "110978.89", "54669.06")
    testDatum(6) = Array("Sub-Saharan Africa", "South Africa", "Beverages", "Offline", "H", "7/10/2012", "482292354", "8/21/2012", "9880", "47.45", "31.79", "468806.00", "314085.20", "154720.80")
    testDatum(7) = Array("Asia", "Laos", "Vegetables", "Online", "L", "2/20/2011", "844532620", "3/20/2011", "4825", "154.06", "90.93", "743339.50", "438737.25", "304602.25")
    testDatum(8) = Array("Asia", "China", "Baby Food", "Online", "C", "4/10/2017", "564251220", "5/12/2017", "3330", "255.28", "159.42", "850082.40", "530868.60", "319213.80")
    testDatum(9) = Array("Sub-Saharan Africa", "Eritrea", "Meat", "Online", "L", "11/21/2014", "411809480", "1/10/2015", "2431", "421.89", "364.69", "1025614.59", "886561.39", "139053.20")

    'Act:
    SUT.Items = testDatum
    actual = SUT.ToCSVString()
    
    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BetterArray_ToCSVString")
Private Sub ToCSVString_1DArray_ValidStringReturned()
    On Error GoTo TestFail

    'Arrange:
    Const expected As String = _
        "Region" & vbCrLf & _
        "Country A" & vbCrLf & _
        "Item Type" & vbCrLf & _
        "Sales Channel" & vbCrLf & _
        "Order Priority" & vbCrLf & _
        "Order Date" & vbCrLf & _
        "Order ID" & vbCrLf & _
        "Ship Date" & vbCrLf & _
        "Units Sold" & vbCrLf & _
        "Unit Price" & vbCrLf & _
        "Unit Cost" & vbCrLf & _
        "Total Revenue" & vbCrLf & _
        "Total Cost" & vbCrLf & _
        "Total Profit"

    Dim testDatum() As Variant
    Dim actual As String

    testDatum = Array("Region", "Country A", "Item Type", "Sales Channel", "Order Priority", "Order Date", "Order ID", "Ship Date", "Units Sold", "Unit Price", "Unit Cost", "Total Revenue", "Total Cost", "Total Profit")

    'Act:
    SUT.Items = testDatum
    actual = SUT.ToCSVString()

    'Assert:
    Assert.AreEqual expected, actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub





