Attribute VB_Name = "TestModule_BetterArray"
Attribute VB_Description = "Unit Tests for 'BetterArray.cls'"
Option Explicit
Option Private Module

'@TestModule
'@Folder("VBABetterArray.Tests")
'@ModuleDescription("Unit Tests for 'BetterArray.cls'")

'@IgnoreModule AssignmentNotUsed, ProcedureNotUsed
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
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Constructor")
Private Sub Constructor_CreatesWithDefaultCapacity_CapacityIsFour()
    On Error GoTo TestFail

    'Arrange:
    Const Expected As Long = 4
    Dim Actual As Long

    'Act:
    Actual = SUT.Capacity

    'Assert:
    Assert.AreEqual Expected, Actual, "Default capacity incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


''''''''''''''''''''''''''''''''''''
' Attribute - DefaultMember - Item '
''''''''''''''''''''''''''''''''''''

'@TestMethod("BetterArray_Items")
Private Sub Items_DefaultMember_DefaultMemberAccessReturnsItems()
    On Error GoTo TestFail

    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim i As Long

    Expected = Gen.GetArray()
    
    'Act:
    For i = LBound(Expected) To UBound(Expected)
        '@Ignore IndexedDefaultMemberAccess
        SUT(i) = Expected(i)
    Next
    Actual = SUT.Items

    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual array does not match expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'''''''''''''''''''
' Prop - Capacity '
'''''''''''''''''''

'@TestMethod("BetterArray_Capacity")
'@Ignore DuplicatedAnnotation
Private Sub Capacity_CanSetCapacity_ReturnedCapacityMatchesSetCapacity()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Long = 20
    Dim Actual As Long
   
    'Act:
    SUT.Capacity = Expected
    Actual = SUT.Capacity

    'Assert:
    Assert.AreEqual Expected, Actual, "Returned capacity does not equal set capacity"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


''''''''''''''''
' Prop - Items '
''''''''''''''''

'@TestMethod("BetterArray_Items")
Private Sub Items_CanAssignOneDimemsionalArray_ReturnedArrayEqualsAssignedArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant

    Expected = Gen.GetArray(AG_VARIANT, AG_ONEDIMENSION)
   
    'Act:
    SUT.Items = Expected
    Actual = SUT.Items

    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual array does not match expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Items")
Private Sub Items_CanAssignMultiDimemsionalArray_ReturnedArrayEqualsAssignedArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    
    Expected = Gen.GetArray(AG_VARIANT, AG_MULTIDIMENSION)
 
    'Act:
    SUT.Items = Expected
    Actual = SUT.Items

    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual array does not match expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Items")
' NOTE: does not use SequenceEquals due to Rubberduck issue: https://github.com/rubberduck-vba/Rubberduck/issues/5161
Private Sub Items_CanAssignJaggedArray_ReturnedArrayEqualsAssignedArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestResult As Boolean

    Expected = Gen.GetArray(AG_VARIANT, AG_JAGGED)
    
    'Act:
    SUT.Items = Expected
    Actual = SUT.Items
    
    TestResult = SequenceEquals_JaggedArray(Expected, Actual)

    'Assert:
    Assert.IsTrue TestResult, "Contents of expected and actual do not match"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'''''''''''''''''
' Prop - Length '
'''''''''''''''''

'@TestMethod("BetterArray_Length")
Private Sub Length_NewUninitArray_EmptyArrayHasLengthZero()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Long
    Dim Actual As Long
    
    Expected = 0
   
    'Act:
    Actual = SUT.Length

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual length <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Length")
Private Sub Length_FromAssignedOneDimensionalArray_ReturnedLengthEqualsOriginalArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim Expected As Long
    Dim Actual As Long
    
    Expected = TEST_ARRAY_LENGTH
    TestArray = Gen.GetArray(AG_VARIANT, AG_ONEDIMENSION)
    
    'Act:
    SUT.Items = TestArray
    Actual = SUT.Length

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual length <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Length")
Private Sub Length_FromAssignedMultiDimensionalArray_ReturnedLengthEqualsOriginalArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim Expected As Long
    Dim Actual As Long

    Expected = TEST_ARRAY_LENGTH
    TestArray = Gen.GetArray(AG_VARIANT, AG_MULTIDIMENSION)

    'Act:
    SUT.Items = TestArray
    Actual = SUT.Length
    
    'Assert:
    Assert.AreEqual Expected, Actual, "Actual length <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Length")
Private Sub Length_FromAssignedJaggedArray_ReturnedLengthEqualsOriginalArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim Expected As Long
    Dim Actual As Long

    Expected = TEST_ARRAY_LENGTH
    TestArray = Gen.GetArray(AG_VARIANT, AG_JAGGED)

    'Act:
    SUT.Items = TestArray
    Actual = SUT.Length

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual length <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Length")
Private Sub Upperbound_FromAssignedOneDimensionalArray_ReturnedUpperBoundEqualsOriginalArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim Expected As Long
    Dim Actual As Long

    TestArray = Gen.GetArray(AG_VARIANT, AG_ONEDIMENSION)
    Expected = UBound(TestArray)

    'Act:
    SUT.Items = TestArray
    Actual = SUT.UpperBound

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual upperbound <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'''''''''''''''
' Prop - LowerBound '
'''''''''''''''


'@TestMethod("BetterArray_LowerBound")
Private Sub LowerBound_FromAssignedOneDimensionalArray_ReturnedLowerBoundEqualsOriginalArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim Expected As Long
    Dim Actual As Long

    TestArray = Gen.GetArray(AG_VARIANT, AG_ONEDIMENSION)
    Expected = LBound(TestArray)
    
    'Act:
    SUT.Items = TestArray
    Actual = SUT.LowerBound

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual LowerBound <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_LowerBound")
Private Sub LowerBound_ChangingLowerBoundOfAssignedArray_ReturnedArrayHasNewLowerBound()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim ReturnedItems As Variant
    Dim Expected As Long
    Dim Actual As Long
    Dim OldLowerBound As Long
    
    TestArray = Gen.GetArray()
    OldLowerBound = LBound(TestArray)
        
    'Act:
    SUT.Items = TestArray
    Expected = OldLowerBound + 1
    SUT.LowerBound = Expected
    ReturnedItems = SUT.Items
    Actual = LBound(ReturnedItems)

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual LowerBound <> expected"
    Assert.AreEqual SUT.LowerBound, Actual, "Actual LowerBound <> SUT.LowerBound prop"
    Assert.AreEqual UBound(TestArray) + 1, UBound(ReturnedItems), "Actual upperbound <> expected"
    Assert.AreEqual SUT.UpperBound, UBound(ReturnedItems), "Actual upperbound <> SUT.UpperBound prop"
    Assert.AreEqual SUT.Length, TEST_ARRAY_LENGTH, "Actual length does not equal expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'''''''''''''''
' Prop - Item '
'''''''''''''''

'@TestMethod("BetterArray_Item")
Private Sub Item_ChangingExistingIndex_ItemIsChanged()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As String = "Hello World"
    Dim TestArray() As Variant
    Dim Actual As Variant
    Dim ActualLowerBound As Long
    Dim ExpectedLowerBound As Long

    TestArray = Gen.GetArray(AG_VARIANT, AG_ONEDIMENSION)
    ExpectedLowerBound = LBound(TestArray)
    
    'Act:
    SUT.Items = TestArray
    SUT.Item(1) = Expected
    Actual = SUT.Item(1)
    ActualLowerBound = SUT.LowerBound

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
    Assert.AreEqual ExpectedLowerBound, ActualLowerBound, "Actual LowerBound does not equal expected LowerBound"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Item")
Private Sub Item_ChangingIndexOverUpperBound_ItemIsPushed()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As String = "Hello World"
    Dim TestArray() As Variant
    Dim Actual As Variant
    Dim ActualLowerBound As Long
    Dim ExpectedLowerBound As Long
    
    TestArray = Gen.GetArray(AG_VARIANT, AG_ONEDIMENSION)
    ExpectedLowerBound = LBound(TestArray)
    
    'Act:
    SUT.Items = TestArray
    SUT.Item(SUT.UpperBound + 1) = Expected
    Actual = SUT.Item(SUT.UpperBound)
    ActualLowerBound = SUT.LowerBound
    
    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
    Assert.AreEqual TEST_ARRAY_LENGTH + 1, SUT.Length, "Actual length does not match expected length"
    Assert.AreEqual ExpectedLowerBound, ActualLowerBound, "Actual LowerBound does not match expected LowerBound"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Item")
Private Sub Item_ChangingIndexBelowLowerBound_ItemIsUnshifted()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As String = "Hello World"
    Dim TestArray() As Variant
    Dim Actual As Variant
    Dim ExpectedLowerBound As Long
    Dim ActualLowerBound As Long

    TestArray = Gen.GetArray(AG_VARIANT, AG_ONEDIMENSION)
    ExpectedLowerBound = LBound(TestArray)
    
    'Act:
    SUT.Items = TestArray
    SUT.Item(SUT.LowerBound - 10) = Expected
    Actual = SUT.Item(SUT.LowerBound)
    ActualLowerBound = SUT.LowerBound

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual result does not match expected result"
    Assert.AreEqual TEST_ARRAY_LENGTH + 1, SUT.Length, "Actual length does not match expected length"
    Assert.AreEqual ExpectedLowerBound, ActualLowerBound, "Actual LowerBound does not match expected LowerBound"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Item")
Private Sub Item_GetScalarValue_ValueReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim Expected As Variant
    Dim Actual As Variant
       
    TestArray = Gen.GetArray(AG_VARIANT, AG_ONEDIMENSION)
    Expected = TestArray(1)
    
    'Act:
    SUT.Items = TestArray
    Actual = SUT.Item(1)

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Item")
Private Sub Item_GetObject_SameObjectReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim Expected As Object
    Dim Actual As Object

    TestArray = Gen.GetArray(AG_OBJECT, AG_ONEDIMENSION)
    Set Expected = TestArray(1)
    
    'Act:
    SUT.Items = TestArray
    Set Actual = SUT.Item(1)

    'Assert:
    Assert.AreSame Expected, Actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
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
    Dim NewType As ArrayTypes
    NewType = ArrayTypes.BA_UNDEFINED
    'Act:
    SUT.ArrayType = NewType

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
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
    Dim StartingType As ArrayTypes
    Dim NewType As ArrayTypes
    Dim Expected() As Variant
    Dim Actual() As Variant
    
    StartingType = ArrayTypes.BA_UNALLOCATED
    NewType = ArrayTypes.BA_UNALLOCATED
    ReDim Expected(SUT.LowerBound)
    
    'Act:
    Assert.AreEqual StartingType, SUT.ArrayType
    SUT.ArrayType = NewType
    Actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Expected undefined to return empty array"
    Assert.AreEqual NewType, SUT.ArrayType, "Actual type <> Expected"
    Assert.AreEqual 0&, SUT.Length, "Actual length <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

' Starting Unalloc - Change to OneDimension
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_UnallocToOneDimension_Success()
    On Error GoTo TestFail
       
    'Arrange:
    Dim StartingType As ArrayTypes
    Dim NewType As ArrayTypes
    Dim Expected() As Variant
    Dim Actual() As Variant
    
    StartingType = ArrayTypes.BA_UNALLOCATED
    NewType = ArrayTypes.BA_ONEDIMENSION
    ReDim Expected(SUT.LowerBound)
    
    'Act:
    Assert.AreEqual StartingType, SUT.ArrayType
    SUT.ArrayType = NewType
    Actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Expected undefined to return empty array"
    Assert.AreEqual NewType, SUT.ArrayType, "Actual type <> Expected"
    Assert.AreEqual 0&, SUT.Length, "Actual length <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


' Starting Unalloc - Change to MultiDimension
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_UnallocToMultiDimension_Success()
    On Error GoTo TestFail
       
    'Arrange:
    Dim StartingType As ArrayTypes
    Dim NewType As ArrayTypes
    Dim Expected() As Variant
    Dim Actual() As Variant
    
    StartingType = ArrayTypes.BA_UNALLOCATED
    NewType = ArrayTypes.BA_MULTIDIMENSION
    ReDim Expected(SUT.LowerBound)
    
    'Act:
    Assert.AreEqual StartingType, SUT.ArrayType
    SUT.ArrayType = NewType
    Actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Expected undefined to return empty array"
    Assert.AreEqual NewType, SUT.ArrayType, "Actual type <> Expected"
    Assert.AreEqual 0&, SUT.Length, "Actual length <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

' Starting Unalloc - Change to Jagged
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_UnallocToJagged_Success()
    On Error GoTo TestFail
       
    'Arrange:
    Dim StartingType As ArrayTypes
    Dim NewType As ArrayTypes
    Dim Expected() As Variant
    Dim Actual() As Variant
    
    StartingType = ArrayTypes.BA_UNALLOCATED
    NewType = ArrayTypes.BA_JAGGED
    ReDim Expected(SUT.LowerBound)
    
    'Act:
    Assert.AreEqual StartingType, SUT.ArrayType
    SUT.ArrayType = NewType
    Actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Expected undefined to return empty array"
    Assert.AreEqual NewType, SUT.ArrayType, "Actual type <> Expected"
    Assert.AreEqual 0&, SUT.Length, "Actual length <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


' Starting OneDimension - Change to Undefined
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_OneDimensionToUndefined_ThrowsError()
    Const ExpectedError As Long = ErrorCodes.EC_CANNOT_CONVERT_TO_REQUESTED_STRUCTURE
    On Error GoTo TestFail
    
    'Arrange:
    Dim StartingType As ArrayTypes
    Dim NewType As ArrayTypes
    Dim TestArray() As Variant
    
    StartingType = ArrayTypes.BA_ONEDIMENSION
    NewType = ArrayTypes.BA_UNDEFINED
    TestArray = Gen.GetArray(ArrayType:=StartingType)
    SUT.Items = TestArray
    
    'Act:
    Assert.AreEqual StartingType, SUT.ArrayType
    SUT.ArrayType = NewType
    Assert.AreEqual NewType, SUT.ArrayType
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
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
    Dim StartingType As ArrayTypes
    Dim NewType As ArrayTypes
    Dim TestArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    
    StartingType = ArrayTypes.BA_ONEDIMENSION
    NewType = ArrayTypes.BA_UNALLOCATED
    ReDim Expected(SUT.LowerBound)
    TestArray = Gen.GetArray(ArrayType:=StartingType)
    SUT.Items = TestArray
    
    'Act:
    Assert.AreEqual StartingType, SUT.ArrayType
    SUT.ArrayType = NewType
    Actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Expected undefined to return empty array"
    Assert.AreEqual NewType, SUT.ArrayType, "Actual type <> Expected"
    Assert.AreEqual 0&, SUT.Length, "Actual length <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

' Starting OneDimension - Change to OneDimension
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_OneDimensionToOneDimension_Success()
    On Error GoTo TestFail
       
    'Arrange:
    Dim StartingType As ArrayTypes
    Dim NewType As ArrayTypes
    Dim TestArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    
    StartingType = ArrayTypes.BA_ONEDIMENSION
    NewType = ArrayTypes.BA_ONEDIMENSION
    TestArray = Gen.GetArray(ArrayType:=StartingType)
    SUT.Items = TestArray
    Expected = TestArray
    
    'Act:
    Assert.AreEqual StartingType, SUT.ArrayType
    SUT.ArrayType = NewType
    Actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Expected undefined to return empty array"
    Assert.AreEqual NewType, SUT.ArrayType, "Actual type <> Expected"
    Assert.AreEqual TEST_ARRAY_LENGTH, SUT.Length, "Actual length <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


' Starting OneDimension - Change to MultiDimension
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_OneDimensionToMultiDimension_Success()
    On Error GoTo TestFail
       
    'Arrange:
    Dim StartingType As ArrayTypes
    Dim NewType As ArrayTypes
    Dim TestArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    
    StartingType = ArrayTypes.BA_ONEDIMENSION
    NewType = ArrayTypes.BA_MULTIDIMENSION
    TestArray = Gen.GetArray(ArrayType:=StartingType)
    SUT.Items = TestArray
    Expected = TestArray
    
    'Act:
    Assert.AreEqual StartingType, SUT.ArrayType
    SUT.ArrayType = NewType
    Actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Expected undefined to return empty array"
    Assert.AreEqual NewType, SUT.ArrayType, "Actual type <> Expected"
    Assert.AreEqual TEST_ARRAY_LENGTH, SUT.Length, "Actual length <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


' Starting OneDimension - Change to Jagged
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_OneDimensionToJagged_Success()
    On Error GoTo TestFail
       
    'Arrange:
    Dim StartingType As ArrayTypes
    Dim NewType As ArrayTypes
    Dim TestArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    
    StartingType = ArrayTypes.BA_ONEDIMENSION
    NewType = ArrayTypes.BA_JAGGED
    TestArray = Gen.GetArray(ArrayType:=StartingType)
    SUT.Items = TestArray
    Expected = TestArray
    
    'Act:
    Assert.AreEqual StartingType, SUT.ArrayType
    SUT.ArrayType = NewType
    Actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Expected undefined to return empty array"
    Assert.AreEqual NewType, SUT.ArrayType, "Actual type <> Expected"
    Assert.AreEqual TEST_ARRAY_LENGTH, SUT.Length, "Actual length <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

' Starting MultiDimension - Change to Undefined
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_MultiDimensionToUndefined_ThrowsError()
    Const ExpectedError As Long = ErrorCodes.EC_CANNOT_CONVERT_TO_REQUESTED_STRUCTURE
    On Error GoTo TestFail
    
    'Arrange:
    Dim StartingType As ArrayTypes
    Dim NewType As ArrayTypes
    Dim TestArray() As Variant
    
    StartingType = ArrayTypes.BA_MULTIDIMENSION
    NewType = ArrayTypes.BA_UNDEFINED
    
    TestArray = Gen.GetArray(ArrayType:=StartingType)
    SUT.Items = TestArray
    
    'Act:
    Assert.AreEqual StartingType, SUT.ArrayType
    SUT.ArrayType = NewType
    Assert.AreEqual NewType, SUT.ArrayType

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
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
    Dim StartingType As ArrayTypes
    Dim NewType As ArrayTypes
    Dim TestArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    
    StartingType = ArrayTypes.BA_MULTIDIMENSION
    NewType = ArrayTypes.BA_UNALLOCATED
    ReDim Expected(SUT.LowerBound)
    TestArray = Gen.GetArray(ArrayType:=StartingType)
    SUT.Items = TestArray
    
    'Act:
    Assert.AreEqual StartingType, SUT.ArrayType
    SUT.ArrayType = NewType
    Actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Expected undefined to return empty array"
    Assert.AreEqual NewType, SUT.ArrayType, "Actual type <> Expected"
    Assert.AreEqual 0&, SUT.Length, "Actual length <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

' Starting MultiDimension - Change to OneDimension
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_MultiDimensionToOneDimension_ThrowsError()
    Const ExpectedError As Long = ErrorCodes.EC_CANNOT_CONVERT_TO_REQUESTED_STRUCTURE
    On Error GoTo TestFail
    
    'Arrange:
    Dim StartingType As ArrayTypes
    Dim NewType As ArrayTypes
    Dim TestArray() As Variant
    
    StartingType = ArrayTypes.BA_MULTIDIMENSION
    NewType = ArrayTypes.BA_ONEDIMENSION
    
    TestArray = Gen.GetArray(ArrayType:=StartingType)
    SUT.Items = TestArray
    
    'Act:
    Assert.AreEqual StartingType, SUT.ArrayType
    SUT.ArrayType = NewType
    Assert.AreEqual NewType, SUT.ArrayType

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
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
    Dim StartingType As ArrayTypes
    Dim NewType As ArrayTypes
    Dim TestArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    
    StartingType = ArrayTypes.BA_MULTIDIMENSION
    NewType = ArrayTypes.BA_MULTIDIMENSION
    TestArray = Gen.GetArray(ArrayType:=StartingType)
    SUT.Items = TestArray
    Expected = TestArray
    
    'Act:
    Assert.AreEqual StartingType, SUT.ArrayType
    SUT.ArrayType = NewType
    Actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Expected undefined to return empty array"
    Assert.AreEqual NewType, SUT.ArrayType, "Actual type <> Expected"
    Assert.AreEqual TEST_ARRAY_LENGTH, SUT.Length, "Actual length <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


' Starting MultiDimension - Change to Jagged
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_MultiDimensionToJagged_Success()
    On Error GoTo TestFail
       
    'Arrange:
    Dim StartingType As ArrayTypes
    Dim NewType As ArrayTypes
    Dim TestArray() As Variant
    Dim Actual() As Variant
    
    StartingType = ArrayTypes.BA_MULTIDIMENSION
    NewType = ArrayTypes.BA_JAGGED
    TestArray = Gen.GetArray(ArrayType:=StartingType)
    SUT.Items = TestArray
    
    'Act:
    Assert.AreEqual StartingType, SUT.ArrayType
    SUT.ArrayType = NewType
    Actual = SUT.Items
    
    'Assert:
    Assert.IsTrue SequenceEqualsMutiVsJagged(TestArray, Actual)
    Assert.AreEqual NewType, SUT.ArrayType, "Actual type <> Expected"
    Assert.AreEqual TEST_ARRAY_LENGTH, SUT.Length, "Actual length <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


' Starting Jagged - Change to Undefined
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_JaggedToUndefined_ThrowsError()
    Const ExpectedError As Long = ErrorCodes.EC_CANNOT_CONVERT_TO_REQUESTED_STRUCTURE
    On Error GoTo TestFail
    
    'Arrange:
    Dim StartingType As ArrayTypes
    Dim NewType As ArrayTypes
    Dim TestArray() As Variant
    
    StartingType = ArrayTypes.BA_JAGGED
    NewType = ArrayTypes.BA_UNDEFINED
    
    TestArray = Gen.GetArray(ArrayType:=StartingType)
    SUT.Items = TestArray
    
    'Act:
    Assert.AreEqual StartingType, SUT.ArrayType
    SUT.ArrayType = NewType
    Assert.AreEqual NewType, SUT.ArrayType

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
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
    Dim StartingType As ArrayTypes
    Dim NewType As ArrayTypes
    Dim TestArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    
    StartingType = ArrayTypes.BA_MULTIDIMENSION
    NewType = ArrayTypes.BA_UNALLOCATED
    ReDim Expected(SUT.LowerBound)
    TestArray = Gen.GetArray(ArrayType:=StartingType)
    SUT.Items = TestArray
    
    'Act:
    Assert.AreEqual StartingType, SUT.ArrayType
    SUT.ArrayType = NewType
    Actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Expected undefined to return empty array"
    Assert.AreEqual NewType, SUT.ArrayType, "Actual type <> Expected"
    Assert.AreEqual 0&, SUT.Length, "Actual length <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

' Starting Jagged - Change to OneDimension
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_JaggedToOneDimension_ThrowsError()
    Const ExpectedError As Long = ErrorCodes.EC_CANNOT_CONVERT_TO_REQUESTED_STRUCTURE
    On Error GoTo TestFail
    
    'Arrange:
    Dim StartingType As ArrayTypes
    Dim NewType As ArrayTypes
    Dim TestArray() As Variant
    
    StartingType = ArrayTypes.BA_JAGGED
    NewType = ArrayTypes.BA_ONEDIMENSION
    
    TestArray = Gen.GetArray(ArrayType:=StartingType)
    SUT.Items = TestArray
    
    'Act:
    Assert.AreEqual StartingType, SUT.ArrayType
    SUT.ArrayType = NewType
    Assert.AreEqual NewType, SUT.ArrayType
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
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
    Dim StartingType As ArrayTypes
    Dim NewType As ArrayTypes
    Dim TestArray() As Variant
    Dim Actual() As Variant
    
    StartingType = ArrayTypes.BA_JAGGED
    NewType = ArrayTypes.BA_MULTIDIMENSION
    TestArray = Gen.GetArray(ArrayType:=StartingType)
    SUT.Items = TestArray
    
    'Act:
    Assert.AreEqual StartingType, SUT.ArrayType
    SUT.ArrayType = NewType
    Actual = SUT.Items
    
    'Assert:
    Assert.IsTrue SequenceEqualsMutiVsJagged(Actual, TestArray)
    Assert.AreEqual NewType, SUT.ArrayType, "Actual type <> Expected"
    Assert.AreEqual TEST_ARRAY_LENGTH, SUT.Length, "Actual length <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

' Starting Jagged - Change to Jagged
'@TestMethod("BetterArray_ArrayType")
Private Sub ArrayType_JaggedToJagged_Success()
    On Error GoTo TestFail
       
    'Arrange:
    Dim StartingType As ArrayTypes
    Dim NewType As ArrayTypes
    Dim TestArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    
    StartingType = ArrayTypes.BA_JAGGED
    NewType = ArrayTypes.BA_JAGGED
    TestArray = Gen.GetArray(ArrayType:=StartingType)
    SUT.Items = TestArray
    Expected = TestArray
    
    'Act:
    Assert.AreEqual StartingType, SUT.ArrayType
    SUT.ArrayType = NewType
    Actual = SUT.Items
    
    'Assert:
    Assert.IsTrue SequenceEquals_JaggedArray(Expected, Actual), "Expected undefined to return empty array"
    Assert.AreEqual NewType, SUT.ArrayType, "Actual type <> Expected"
    Assert.AreEqual TEST_ARRAY_LENGTH, SUT.Length, "Actual length <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'''''''''''''''''
' Method - Push '
'''''''''''''''''

'@TestMethod("BetterArray_Push")
Private Sub Push_AddToNewBetterArray_ItemAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As String = "Hello World"
    Const ExpectedLength As Long = 1
    Const ExpectedUpperBound As Long = 0

    Dim Actual As String
    Dim ActualLength As Long
    Dim ActualUpperBound As Long
    
    'Act:
    SUT.Push Expected
    Actual = SUT.Item(SUT.LowerBound)
    ActualLength = SUT.Length
    ActualUpperBound = SUT.UpperBound

    'Assert:
    Assert.AreEqual Expected, Actual, "actual <> expected"
    Assert.AreEqual ExpectedLength, ActualLength, "Length value incorrect"
    Assert.AreEqual ExpectedUpperBound, ActualUpperBound, "Upperbound value incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Push")
Private Sub Push_AddToExistingOneDimensionalArray_ItemAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As String = "Hello World"

    Dim TestArray() As Variant
    Dim Actual As String
    
    TestArray = Gen.GetArray(AG_VARIANT, AG_ONEDIMENSION)
    
    'Act:
    SUT.Items = TestArray
    SUT.Push Expected
    Actual = SUT.Item(SUT.UpperBound)

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
    Assert.AreEqual TEST_ARRAY_LENGTH + 1, SUT.Length, "Length value incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Push")
Private Sub Push_AddToExistingMultidimensionalArray_ItemAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Variant
    Dim Actual As Variant
    Dim TestArray() As Variant
    Dim ReturnedArray() As Variant

    Expected = "Hello World"
    TestArray = Gen.GetArray(AG_VARIANT, AG_MULTIDIMENSION)
    
    'Act:
    SUT.Items = TestArray
    SUT.Push Expected
    ReturnedArray = SUT.Items
    Actual = ReturnedArray(UBound(ReturnedArray), LBound(ReturnedArray, 2))

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
    Assert.AreEqual TEST_ARRAY_LENGTH + 1, SUT.Length, "Length value incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Push")
Private Sub Push_AddToExistingJaggedArray_ItemAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Variant
    Dim Actual As Variant
    Dim TestArray() As Variant
    Dim ReturnedArray() As Variant

    Expected = Gen.GetArray(AG_VARIANT, AG_ONEDIMENSION)
    TestArray = Gen.GetArray(AG_VARIANT, AG_JAGGED)
    
    'Act:
    SUT.Items = TestArray
    SUT.Push Expected
    ReturnedArray = SUT.Items
    Actual = ReturnedArray(UBound(ReturnedArray))
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Element value incorrect"
    Assert.AreEqual TEST_ARRAY_LENGTH + 1, SUT.Length, "Length value incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Push")
Private Sub Push_AddMultipleToNewBetterArray_ItemsAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Long = 1
    Const ExpectedLength As Long = 3
    Const ExpectedUpperBound As Long = 2

    Dim Actual As Long
    Dim ActualLength As Long
    Dim ActualUpperBound As Long
        
    'Act:
    SUT.Push Expected, 2, 3
    Actual = SUT.Item(SUT.LowerBound)
    ActualLength = SUT.Length
    ActualUpperBound = SUT.UpperBound

    'Assert:
    Assert.AreEqual Expected, Actual, "Element value incorrect"
    Assert.AreEqual ExpectedLength, ActualLength, "Length value incorrect"
    Assert.AreEqual ExpectedUpperBound, ActualUpperBound, "Upperbound value incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


''''''''''''''''
' Method - Pop '
''''''''''''''''

'@TestMethod("BetterArray_Pop")
Private Sub Pop_ItemsRemovedByPopUntilEmpty_EmptyArrayHasLengthZero()
    ' Added for coverage of issue #15 - https://github.com/Senipah/VBA-Better-Array/issues/15
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Long
    Dim Actual As Long
    
    Expected = 0
    SUT.Push 1, 2
    
    'Act:
    SUT.Pop
    SUT.Pop
    Actual = SUT.Length

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual length <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Pop")
Private Sub Pop_OneDimensionalArray_LastItemRemoved()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim ActualLowerBound As Long
    Dim ExpectedLowerBound As Long
    Dim Expected As String
    Dim Actual As String
    
    TestArray = Gen.GetArray(AG_STRING, AG_ONEDIMENSION)
    Expected = TestArray(UBound(TestArray))
    ExpectedLowerBound = SUT.LowerBound
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.Pop
    ActualLowerBound = SUT.LowerBound

    'Assert:
    Assert.AreEqual Expected, Actual, "Element value incorrect"
    Assert.AreEqual TEST_ARRAY_LENGTH - 1, SUT.Length, "Length value incorrect"
    Assert.AreEqual UBound(TestArray) - 1, SUT.UpperBound, "Upperbound value incorrect"
    Assert.AreEqual ExpectedLowerBound, ActualLowerBound, "LowerBound value incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Pop")
Private Sub Pop_ArrayLengthIsZero_ReturnsEmpty()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Variant
    Dim ExpectedLowerBound As Long
    Dim ExpectedLength As Long
    Dim ExpectedUpperBound As Long
    Dim Actual As Variant
    Dim ActualLowerBound As Long
    Dim ActualLength As Long
    Dim ActualUpperBound As Long
    
    Expected = Empty
    ExpectedLowerBound = 0
    ExpectedLength = 0
    ExpectedUpperBound = -1
    
    'Act:
    Actual = SUT.Pop
    ActualLowerBound = SUT.LowerBound
    ActualLength = SUT.Length
    ActualUpperBound = SUT.UpperBound

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
    Assert.AreEqual ExpectedLowerBound, ActualLowerBound, "LowerBound value incorrect"
    Assert.AreEqual ExpectedLength, ActualLength, "Length value incorrect"
    Assert.AreEqual ExpectedUpperBound, ActualUpperBound, "Upperbound value incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


''''''''''''''''''
' Method - Shift '
''''''''''''''''''

'@TestMethod("BetterArray_Shift")
Private Sub Shift_ItemsRemovedByShiftUntilEmpty_EmptyArrayHasLengthZero()
    ' Agged for coverage of issue #15 - https://github.com/Senipah/VBA-Better-Array/issues/15
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Long
    Dim Actual As Long
    
    Expected = 0
    SUT.Push 1, 2
    
    'Act:
    SUT.Shift
    SUT.Shift
    Actual = SUT.Length

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual length <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Shift")
Private Sub Shift_OneDimensionalArray_FirstItemRemoved()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim ActualLowerBound As Long
    Dim ExpectedLowerBound As Long
    Dim Expected As String
    Dim Actual As String

    TestArray = Gen.GetArray(AG_STRING, AG_ONEDIMENSION)
    Expected = TestArray(LBound(TestArray))
    ExpectedLowerBound = SUT.LowerBound
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.Shift
    ActualLowerBound = SUT.LowerBound

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
    Assert.AreEqual TEST_ARRAY_LENGTH - 1, SUT.Length, "Length value incorrect"
    Assert.AreEqual UBound(TestArray) - 1, SUT.UpperBound, "Upperbound value incorrect"
    Assert.AreEqual ExpectedLowerBound, ActualLowerBound, "LowerBound value incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Shift")
Private Sub Shift_ArrayLengthIsZero_ReturnsEmpty()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Variant = Empty
    Const ExpectedLowerBound As Long = 0
    Const ExpectedLength As Long = 0
    Const ExpectedUpperBound As Long = -1

    Dim Actual As Variant
    Dim ActualLowerBound As Long
    Dim ActualLength As Long
    Dim ActualUpperBound As Long
    
    'Act:
    Actual = SUT.Shift
    ActualLowerBound = SUT.LowerBound
    ActualLength = SUT.Length
    ActualUpperBound = SUT.UpperBound

    'Assert:
    Assert.AreEqual Expected, Actual, "Element value incorrect"
    Assert.AreEqual ExpectedLowerBound, ActualLowerBound, "LowerBound value incorrect"
    Assert.AreEqual ExpectedLength, ActualLength, "Length value incorrect"
    Assert.AreEqual ExpectedUpperBound, ActualUpperBound, "Upperbound value incorrect"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

''''''''''''''''''''
' Method - Unshift '
''''''''''''''''''''

'@TestMethod("BetterArray_Unshift")
Private Sub Unshift_OneDimensionalArray_ItemAddedToBeginning()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim Expected As String
    Dim Actual As String
    Dim ActualLowerBound As Long
    Dim ExpectedLowerBound As Long
    Dim TestElement As String
    
    TestArray = Gen.GetArray(AG_STRING, AG_ONEDIMENSION)
    TestElement = "Hello World"
    ExpectedLowerBound = SUT.LowerBound
    Expected = TEST_ARRAY_LENGTH + 1
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.Unshift(TestElement)
    ActualLowerBound = SUT.LowerBound

    'Assert:
    Assert.AreEqual Expected, Actual, "Return value incorrect"
    Assert.AreEqual (UBound(TestArray) + 1), SUT.UpperBound, "Upperbound value incorrect"
    Assert.AreEqual ExpectedLowerBound, ActualLowerBound, "LowerBound value incorrect"
    Assert.AreEqual TestElement, SUT.Item(SUT.LowerBound), "Element not inserted at correct position"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Unshift")
Private Sub Unshift_ArrayLengthIsZero_ItemIsPushedToEmptyArray()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Long = 1
    Const ExpectedLowerBound As Long = 0
    Const ExpectedUpperBound As Long = 0
    Const ExpectedElement As String = "Hello World"

    Dim Actual As Long
    Dim ActualLowerBound As Long
    Dim ActualUpperBound As Long
    Dim ActualElement As String
    
    'Act:
    Actual = SUT.Unshift(ExpectedElement)
    ActualLowerBound = SUT.LowerBound
    ActualUpperBound = SUT.UpperBound
    ActualElement = SUT.Item(SUT.LowerBound)

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual length <> expected length"
    Assert.AreEqual ExpectedUpperBound, ActualUpperBound, "Upperbound value incorrect"
    Assert.AreEqual ExpectedLowerBound, ActualLowerBound, "LowerBound value incorrect"
    Assert.AreEqual ExpectedElement, ActualElement, "Actual element <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Unshift")
Private Sub Unshift_MultidimensionalArray_ItemAddedToBeginning()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Long = TEST_ARRAY_LENGTH + 1
    Const ExpectedLowerBound As Long = 0
    Const ExpectedUpperBound As Long = TEST_ARRAY_LENGTH
    Const ExpectedElement As String = "Hello World"

    Dim Actual As Long
    Dim ActualLowerBound As Long
    Dim ActualUpperBound As Long
    Dim ActualElement As String
    Dim TestArray() As Variant
    Dim ReturnedItems() As Variant

    TestArray = Gen.GetArray(AG_VARIANT, AG_MULTIDIMENSION)
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.Unshift(ExpectedElement)
    ReturnedItems = SUT.Items
    ActualLowerBound = SUT.LowerBound
    ActualUpperBound = SUT.UpperBound
    ActualElement = ReturnedItems(LBound(ReturnedItems), LBound(ReturnedItems, 2))

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual length <> expected length"
    Assert.AreEqual ExpectedUpperBound, ActualUpperBound, "Upperbound value incorrect"
    Assert.AreEqual ExpectedLowerBound, ActualLowerBound, "LowerBound value incorrect"
    Assert.AreEqual ExpectedElement, ActualElement, "Actual element <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'''''''''''''''''''
' Method - Concat '
'''''''''''''''''''
'@TestMethod("BetterArray_Concat")
Private Sub Concat_AddOneDimArrayToOneDimIssue7Coverage_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim ExpectedLength As Long
    Dim ActualLength As Long
    Dim ExpectedUpperBound As Long
    Dim ActualUpperBound As Long
    
    ExpectedLength = TEST_ARRAY_LENGTH
    Expected = Gen.GetArray(Length:=ExpectedLength)
    ExpectedUpperBound = UBound(Expected)
    
    Dim FirstArray(0 To 6) As Variant
    Dim SecondArray(0 To TEST_ARRAY_LENGTH - (7 + 1)) As Variant
    Dim i As Long
    For i = LBound(Expected) To 6
        FirstArray(i) = Expected(i)
    Next
    For i = 7 To UBound(Expected)
        SecondArray(i - 7) = Expected(i)
    Next
    
    For i = LBound(FirstArray) To UBound(FirstArray)
        SUT.Push FirstArray(i)
    Next
    
    'Act:
    Actual = SUT.Concat(SecondArray).Items
    ActualLength = SUT.Length
    ActualUpperBound = SUT.UpperBound
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> Expected"
    Assert.AreEqual ExpectedLength, ActualLength, "Actual length <> expected"
    Assert.AreEqual ExpectedUpperBound, ActualUpperBound, "Actual UpperBound <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Concat")
Private Sub Concat_AddOneDimArrayToEmptyInternal_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim ExpectedLength As Long
    Dim ActualLength As Long
    Dim ExpectedUpperBound As Long
    Dim ActualUpperBound As Long
    
    ExpectedLength = TEST_ARRAY_LENGTH
    Expected = Gen.GetArray(Length:=ExpectedLength)
    ExpectedUpperBound = UBound(Expected)
    
    'Act:
    Actual = SUT.Concat(Expected).Items
    ActualLength = SUT.Length
    ActualUpperBound = SUT.UpperBound
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> Expected"
    Assert.AreEqual ExpectedLength, ActualLength, "Actual length <> expected"
    Assert.AreEqual ExpectedUpperBound, ActualUpperBound, "Actual UpperBound <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Concat")
Private Sub Concat_AddMultipleOneDimArrayToEmptyInternal_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim ExpectedLength As Long
    Dim ActualLength As Long
    Dim ExpectedUpperBound As Long
    Dim ActualUpperBound As Long
    Dim FirstAray() As Variant
    Dim SecondArray() As Variant
    
    FirstAray = Array(1, 2, 3)
    SecondArray = Array(4, 5, 6)
    Expected = Array(1, 2, 3, 4, 5, 6)
    ExpectedLength = 6
    ExpectedUpperBound = UBound(Expected)
    
    'Act:
    Actual = SUT.Concat(FirstAray, SecondArray).Items
    ActualLength = SUT.Length
    ActualUpperBound = SUT.UpperBound
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> Expected"
    Assert.AreEqual ExpectedLength, ActualLength, "Actual length <> expected"
    Assert.AreEqual ExpectedUpperBound, ActualUpperBound, "Actual UpperBound <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Concat")
Private Sub Concat_AddOneDimArrayToExistingOneDimArray_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim ExpectedLength As Long
    Dim ActualLength As Long
    Dim ExpectedUpperBound As Long
    Dim ActualUpperBound As Long
    Dim FirstArray() As Variant
    Dim SecondArray() As Variant
    
    FirstArray = Gen.GetArray()
    SecondArray = Gen.GetArray()
    Expected = Gen.ConcatArraysOfSameStructure(AG_ONEDIMENSION, FirstArray, SecondArray)
    ExpectedLength = Gen.GetArrayLength(Expected)
    ExpectedUpperBound = UBound(Expected)
    
    'Act:
    SUT.Items = FirstArray
    Actual = SUT.Concat(SecondArray).Items
    ActualLength = SUT.Length
    ActualUpperBound = SUT.UpperBound
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> Expected"
    Assert.AreEqual ExpectedLength, ActualLength, "Actual length <> expected"
    Assert.AreEqual ExpectedUpperBound, ActualUpperBound, "Actual UpperBound <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Concat")
Private Sub Concat_AddMultiDimArrayToEmptyInternal_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim ExpectedLength As Long
    Dim ActualLength As Long
    Dim ExpectedUpperBound As Long
    Dim ActualUpperBound As Long
    
    ExpectedLength = TEST_ARRAY_LENGTH
    Expected = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION, Length:=ExpectedLength)
    ExpectedUpperBound = UBound(Expected)
    
    'Act:
    Actual = SUT.Concat(Expected).Items
    ActualLength = SUT.Length
    ActualUpperBound = SUT.UpperBound
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> Expected"
    Assert.AreEqual ExpectedLength, ActualLength, "Actual length <> expected"
    Assert.AreEqual ExpectedUpperBound, ActualUpperBound, "Actual UpperBound <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Concat")
Private Sub Concat_AddMultiDimArrayToExistingMultiDimArray_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim FirstArray() As Variant
    Dim SecondArray() As Variant
    Dim Actual() As Variant
    Dim ExpectedLength As Long
    Dim ActualLength As Long
    Dim ExpectedUpperBound As Long
    Dim ActualUpperBound As Long
    
    ExpectedLength = TEST_ARRAY_LENGTH * 2
    FirstArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SecondArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    Expected = Gen.ConcatArraysOfSameStructure(AG_MULTIDIMENSION, FirstArray, SecondArray)
    ExpectedUpperBound = UBound(Expected)
    
    'Act:
    SUT.Items = FirstArray
    Actual = SUT.Concat(SecondArray).Items
    ActualLength = SUT.Length
    ActualUpperBound = SUT.UpperBound
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> Expected"
    Assert.AreEqual ExpectedLength, ActualLength, "Actual length <> expected"
    Assert.AreEqual ExpectedUpperBound, ActualUpperBound, "Actual UpperBound <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Concat")
Private Sub Concat_AddJaggedArrayToEmptyInternal_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim ExpectedLength As Long
    Dim ActualLength As Long
    Dim ExpectedUpperBound As Long
    Dim ActualUpperBound As Long
    Dim TestResult As Boolean
    
    Expected = Gen.GetArray(ArrayType:=AG_JAGGED)
    ExpectedLength = Gen.GetArrayLength(Expected)
    ExpectedUpperBound = UBound(Expected)
    
    'Act:
    Actual = SUT.Concat(Expected).Items
    ActualLength = SUT.Length
    ActualUpperBound = SUT.UpperBound
    TestResult = SequenceEquals_JaggedArray(Expected, Actual)
    
    'Assert:
    Assert.IsTrue TestResult, "Actual <> Expected"
    Assert.AreEqual ExpectedLength, ActualLength, "Actual length <> expected"
    Assert.AreEqual ExpectedUpperBound, ActualUpperBound, "Actual UpperBound <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Concat")
Private Sub Concat_AddJaggedArrayToExistingJagged_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim FirstArray() As Variant
    Dim SecondArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim ExpectedLength As Long
    Dim ActualLength As Long
    Dim ExpectedUpperBound As Long
    Dim ActualUpperBound As Long
    Dim TestResult As Boolean
    
    FirstArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    SecondArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    
    Expected = Gen.ConcatArraysOfSameStructure(AG_JAGGED, FirstArray, SecondArray)
    ExpectedLength = Gen.GetArrayLength(Expected)
    ExpectedUpperBound = UBound(Expected)
    
    'Act:
    SUT.Items = FirstArray
    Actual = SUT.Concat(SecondArray).Items
    ActualLength = SUT.Length
    ActualUpperBound = SUT.UpperBound
    TestResult = SequenceEquals_JaggedArray(Expected, Actual)
    
    'Assert:
    Assert.IsTrue TestResult, "Actual <> Expected"
    Assert.AreEqual ExpectedLength, ActualLength, "Actual length <> expected"
    Assert.AreEqual ExpectedUpperBound, ActualUpperBound, "Actual UpperBound <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Concat")
Private Sub Concat_AddOneDimArrayToExistingJagged_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim FirstArray() As Variant
    Dim SecondArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim ExpectedLength As Long
    Dim ActualLength As Long
    Dim ExpectedUpperBound As Long
    Dim ActualUpperBound As Long
    Dim TestResult As Boolean
    
    FirstArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    SecondArray = Gen.GetArray(ArrayType:=AG_ONEDIMENSION)
    
    Expected = Gen.ConcatArraysOfSameStructure(AG_JAGGED, FirstArray, SecondArray)
    ExpectedLength = Gen.GetArrayLength(Expected)
    ExpectedUpperBound = UBound(Expected)
    
    'Act:
    SUT.Items = FirstArray
    Actual = SUT.Concat(SecondArray).Items
    ActualLength = SUT.Length
    ActualUpperBound = SUT.UpperBound
    TestResult = SequenceEquals_JaggedArray(Expected, Actual)
    
    'Assert:
    Assert.IsTrue TestResult, "Actual <> Expected"
    Assert.AreEqual ExpectedLength, ActualLength, "Actual length <> expected"
    Assert.AreEqual ExpectedUpperBound, ActualUpperBound, "Actual UpperBound <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Concat")
Private Sub Concat_AddOneDimArrayToExistingMulti_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim FirstArray() As Variant
    Dim SecondArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim ExpectedLength As Long
    Dim ActualLength As Long
    Dim ExpectedUpperBound As Long
    Dim ActualUpperBound As Long
    
    ReDim FirstArray(1 To 2, 1 To 2)
    FirstArray(1, 1) = "Foo"
    FirstArray(1, 2) = "Bar"
    FirstArray(2, 1) = "Fizz"
    FirstArray(2, 2) = "Buzz"
    
    SecondArray = Array(1, 2, 3)
    
    ReDim Expected(1 To 5, 1 To 2)
    Expected(1, 1) = FirstArray(1, 1)
    Expected(1, 2) = FirstArray(1, 2)
    Expected(2, 1) = FirstArray(2, 1)
    Expected(2, 2) = FirstArray(2, 2)
    Expected(3, 1) = SecondArray(0)
    Expected(3, 2) = Empty
    Expected(4, 1) = SecondArray(1)
    Expected(4, 2) = Empty
    Expected(5, 1) = SecondArray(2)
    Expected(5, 2) = Empty
    
    ExpectedLength = 5
    ExpectedUpperBound = 5
    
    'Act:
    SUT.Items = FirstArray
    Actual = SUT.Concat(SecondArray).Items
    ActualLength = SUT.Length
    ActualUpperBound = SUT.UpperBound
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> Expected"
    Assert.AreEqual ExpectedLength, ActualLength, "Actual length <> expected"
    Assert.AreEqual ExpectedUpperBound, ActualUpperBound, "Actual UpperBound <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    
End Sub

'@TestMethod("BetterArray_Concat")
Private Sub Concat_AddMultiDimArrayToExistingOneDimArray_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim FirstArray() As Variant
    Dim SecondArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim ExpectedLength As Long
    Dim ActualLength As Long
    Dim ExpectedUpperBound As Long
    Dim ActualUpperBound As Long
    
    FirstArray = Array(1, 2, 3)
    ReDim SecondArray(1 To 2, 1 To 2)
    SecondArray(1, 1) = "Foo"
    SecondArray(1, 2) = "Bar"
    SecondArray(2, 1) = "Fizz"
    SecondArray(2, 2) = "Buzz"
    
    ReDim Expected(0 To 4, 0 To 1)
    Expected(0, 0) = FirstArray(0)
    Expected(0, 1) = Empty
    Expected(1, 0) = FirstArray(1)
    Expected(1, 1) = Empty
    Expected(2, 0) = FirstArray(2)
    Expected(2, 1) = Empty
    Expected(3, 0) = SecondArray(1, 1)
    Expected(3, 1) = SecondArray(1, 2)
    Expected(4, 0) = SecondArray(2, 1)
    Expected(4, 1) = SecondArray(2, 2)
    
    ExpectedLength = 5
    ExpectedUpperBound = 4
    
    'Act:
    SUT.Items = FirstArray
    Actual = SUT.Concat(SecondArray).Items
    ActualLength = SUT.Length
    ActualUpperBound = SUT.UpperBound
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> Expected"
    Assert.AreEqual ExpectedLength, ActualLength, "Actual length <> expected"
    Assert.AreEqual ExpectedUpperBound, ActualUpperBound, "Actual UpperBound <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Concat")
Private Sub Concat_AddMultiDimArrayDepth3ToExistingOneDimArray_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim FirstArray() As Variant
    Dim SecondArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim ExpectedLength As Long
    Dim ActualLength As Long
    Dim ExpectedUpperBound As Long
    Dim ActualUpperBound As Long
    
    FirstArray = Array(1, 2, 3)
    ReDim SecondArray(1 To 2, 1 To 2, 1 To 2)
    SecondArray(1, 1, 1) = "Foo"
    SecondArray(1, 1, 2) = "Bar"
    SecondArray(1, 2, 1) = "Fizz"
    SecondArray(1, 2, 2) = "Buzz"
    SecondArray(2, 1, 1) = "Foo"
    SecondArray(2, 1, 2) = "Bar"
    SecondArray(2, 2, 1) = "Fizz"
    SecondArray(2, 2, 2) = "Buzz"
    
    ReDim Expected(0 To 4, 0 To 1, 0 To 1)
    Expected(0, 0, 0) = FirstArray(0)
    Expected(1, 0, 0) = FirstArray(1)
    Expected(2, 0, 0) = FirstArray(2)
    
    Expected(3, 0, 0) = SecondArray(1, 1, 1)
    Expected(3, 0, 1) = SecondArray(1, 1, 2)
    Expected(3, 1, 0) = SecondArray(1, 2, 1)
    Expected(3, 1, 1) = SecondArray(1, 2, 2)
    
    Expected(4, 0, 0) = SecondArray(2, 1, 1)
    Expected(4, 0, 1) = SecondArray(2, 1, 2)
    Expected(4, 1, 0) = SecondArray(2, 2, 1)
    Expected(4, 1, 1) = SecondArray(2, 2, 2)
    
    ExpectedLength = 5
    ExpectedUpperBound = 4
    
    'Act:
    SUT.Items = FirstArray
    Actual = SUT.Concat(SecondArray).Items
    ActualLength = SUT.Length
    ActualUpperBound = SUT.UpperBound
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> Expected"
    Assert.AreEqual ExpectedLength, ActualLength, "Actual length <> expected"
    Assert.AreEqual ExpectedUpperBound, ActualUpperBound, "Actual UpperBound <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Concat")
Private Sub Concat_AddJaggedArrayToExistingOneDimArray_SuccessAdded()
    On Error GoTo TestFail
    
    'Arrange:
    Dim FirstArray() As Variant
    Dim SecondArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim ExpectedLength As Long
    Dim ActualLength As Long
    Dim ExpectedUpperBound As Long
    Dim ActualUpperBound As Long
    Dim TestResult As Boolean
    
    FirstArray = Gen.GetArray(ArrayType:=AG_ONEDIMENSION)
    SecondArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    
    Expected = Gen.ConcatArraysOfSameStructure(AG_JAGGED, FirstArray, SecondArray)
    ExpectedLength = Gen.GetArrayLength(Expected)
    ExpectedUpperBound = UBound(Expected)
    
    'Act:
    SUT.Items = FirstArray
    Actual = SUT.Concat(SecondArray).Items
    ActualLength = SUT.Length
    ActualUpperBound = SUT.UpperBound
    TestResult = SequenceEquals_JaggedArray(Expected, Actual)
    
    'Assert:
    Assert.IsTrue TestResult, "Actual <> Expected"
    Assert.AreEqual ExpectedLength, ActualLength, "Actual length <> expected"
    Assert.AreEqual ExpectedUpperBound, ActualUpperBound, "Actual UpperBound <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Concat")
Private Sub Concat_AddEmptyToEmpty_ReturnsEmptyArrayWith1Slot()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    
    'Act:
    SUT.Concat Expected
    ReDim Expected(SUT.LowerBound)
    Actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, (SUT.LowerBound = 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'''''''''''''''''''''''''''''''
' Method - CopyFromCollection '
'''''''''''''''''''''''''''''''

'@TestMethod("BetterArray_CopyFromCollection")
Private Sub CopyFromCollection_AddCollectionToEmpty_CollectionConverted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestCollection As Collection
    Dim i As Long
    
    Expected = Gen.GetArray
    Set TestCollection = New Collection
    For i = LBound(Expected) To UBound(Expected)
        TestCollection.Add Expected(i)
    Next
    
    'Act:
    Actual = SUT.CopyFromCollection(TestCollection).Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_CopyFromCollection")
Private Sub CopyFromCollection_AddCollectionToExistingOneDimArray_ArrayReplacedWithCollectionValues()
    On Error GoTo TestFail
    
    'Arrange:
    Dim InitialArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestCollection As Collection
    Dim i As Long
    
    InitialArray = Gen.GetArray
    Expected = Gen.GetArray
    Set TestCollection = New Collection
    For i = LBound(Expected) To UBound(Expected)
        TestCollection.Add Expected(i)
    Next
    SUT.Items = InitialArray
    'Act:
    Actual = SUT.CopyFromCollection(TestCollection).Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'''''''''''''''''''''
' Method - ToString '
'''''''''''''''''''''

'@TestMethod("BetterArray_ToString")
Private Sub ToString_FromOneDimArray_CorrectStringRepresentationReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As String = "{1,2,3}"
    Dim Actual As String
    Dim TestArray() As Variant
    TestArray = Array(1, 2, 3)
    'Act:
    SUT.Items = TestArray
    Actual = SUT.ToString()

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_ToString")
Private Sub ToString_FromOneDimArrayPrettyPrint_CorrectStringRepresentationReturned()
    On Error GoTo TestFail
    
    Const Expected As String = "{1, 2, 3}"
    Dim Actual As String
    Dim TestArray() As Variant
    TestArray = Array(1, 2, 3)
    'Act:
    SUT.Items = TestArray
    Actual = SUT.ToString(PrettyPrint:=True)

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_ToString")
Private Sub ToString_FromOneDimArrayCustomDelimiters_CorrectStringRepresentationReturned()
    On Error GoTo TestFail
    
    Const Expected As String = "[1,2,3]"
    Dim Actual As String
    Dim TestArray() As Variant
    TestArray = Array(1, 2, 3)
    'Act:
    SUT.Items = TestArray
    Actual = SUT.ToString(OpeningDelimiter:="[", ClosingDelimiter:="]")
    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_ToString")
Private Sub ToString_FromJaggedArray_CorrectStringRepresentationReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As String = "{{1,2},{3,4}}"
    Dim Actual As String
    Dim TestArray() As Variant
    
    TestArray = Array(Array(1, 2), Array(3, 4))
    'Act:
    SUT.Items = TestArray
    Actual = SUT.ToString()

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_ToString")
Private Sub ToString_FromJaggedArrayPrettyPrint_CorrectStringRepresentationReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As String = "{" & vbCrLf _
                             & "  {1, 2}, " & vbCrLf _
                             & "  {3, 4}" & vbCrLf _
                             & "}"
    Dim Actual As String
    Dim TestArray() As Variant
    
    TestArray = Array(Array(1, 2), Array(3, 4))
    'Act:
    SUT.Items = TestArray
    Actual = SUT.ToString(True)

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_ToString")
Private Sub ToString_FromEmptyArray_CorrectStringRepresentationReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As String = "{}"
    Dim Actual As String
    Dim TestArray() As Variant
    
    'Act:
    SUT.Items = TestArray
    Actual = SUT.ToString()

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_ToString")
Private Sub ToString_FromEmptyArrayPrettyPrint_CorrectStringRepresentationReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As String = "{}"
    Dim Actual As String
    Dim TestArray() As Variant
    
    'Act:
    SUT.Items = TestArray
    Actual = SUT.ToString()

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'''''''''''''''''''''
' Method - IsSorted '
'''''''''''''''''''''

'@TestMethod("BetterArray_IsSorted")
Private Sub IsSorted_SortedOneDimArray_ReturnsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Boolean
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    Expected = True
    TestArray = Array(1, 2, 3)
    SUT.Items = TestArray
    'Act:
    Actual = SUT.IsSorted

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_IsSorted")
Private Sub IsSorted_UnsortedOneDimArray_ReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Boolean
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    Expected = False
    TestArray = Array(2, 1, 3)
    SUT.Items = TestArray
    'Act:
    Actual = SUT.IsSorted

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_IsSorted")
Private Sub IsSorted_SortedMultiDimArray_ReturnsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Boolean
    Dim Actual As Boolean
    Dim TestArray(0 To 1, 0 To 1) As Variant
    
    Expected = False
    TestArray(0, 0) = "Foo"
    TestArray(0, 1) = 1
    TestArray(1, 0) = "Bar"
    TestArray(1, 1) = 2
    SUT.Items = TestArray
    'Act:
    Actual = SUT.IsSorted

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_IsSorted")
Private Sub IsSorted_UnsortedMultiDimArray_ReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Boolean
    Dim Actual As Boolean
    Dim TestArray(0 To 1, 0 To 1) As Variant
    
    Expected = False
    TestArray(0, 0) = "Foo"
    TestArray(0, 1) = 2
    TestArray(1, 0) = "Bar"
    TestArray(1, 1) = 1
    SUT.Items = TestArray
    'Act:
    Actual = SUT.IsSorted

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_IsSorted")
Private Sub IsSorted_SortedJaggedArray_ReturnsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Boolean
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    Expected = True
    TestArray = Array(Array("Foo", 1), Array("Bar", 1))
    SUT.Items = TestArray
    'Act:
    Actual = SUT.IsSorted(1)

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_IsSorted")
Private Sub IsSorted_UnsortedJaggedArray_ReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Boolean
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    Expected = False
    TestArray = Array(Array("Foo", 2), Array("Bar", 1))
    SUT.Items = TestArray
    'Act:
    Actual = SUT.IsSorted(1)

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_IsSorted")
Private Sub IsSorted_EmptyArray_ReturnsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Boolean
    Dim Actual As Boolean
    
    Expected = True
    'Act:
    Actual = SUT.IsSorted

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> Expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_IsSorted")
Private Sub IsSorted_JaggedArrayWithMoreThan2Dimensions_RaisesError()
    Const ExpectedError As Long = ErrorCodes.EC_EXCEEDS_MAX_SORT_DEPTH
    On Error GoTo TestFail

    'Arrange
    Dim TestArray() As Variant
    '@Ignore VariableNotUsed
    Dim Actual As Boolean
    TestArray = Gen.GetArray(ArrayType:=AG_JAGGED, Depth:=3)
    SUT.Items = TestArray
    'Act
    Actual = SUT.IsSorted

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
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
    Dim TestArray() As Variant
    
    TestArray = Gen.GetArray(AG_OBJECT)
    SUT.Items = TestArray
    'Act:
    SUT.Sort
    
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("BetterArray_Sort")
Private Sub Sort_OneDimArrayQuicksortRecursive_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    TestArray = Gen.GetArray()
    SUT.Items = TestArray
    SUT.SortMethod = SM_QUICKSORT_RECURSIVE
    'Act:
    SUT.Sort
    Actual = SUT.IsSorted
    
    'Assert:
    Assert.IsTrue Actual, "Array not sorted"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Sort")
Private Sub Sort_OneDimArrayQuicksortRecursiveNegativeBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    TestArray = Gen.GetArray()
    SUT.LowerBound = -10
    SUT.Items = TestArray
    SUT.SortMethod = SM_QUICKSORT_RECURSIVE
    'Act:
    SUT.Sort
    Actual = SUT.IsSorted
    
    'Assert:
    Assert.IsTrue Actual, "Array not sorted"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Sort")
Private Sub Sort_OneDimArrayQuicksortRecursivePositiveBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    TestArray = Gen.GetArray()
    SUT.LowerBound = 10
    SUT.Items = TestArray
    SUT.SortMethod = SM_QUICKSORT_RECURSIVE
    'Act:
    SUT.Sort
    Actual = SUT.IsSorted
    
    'Assert:
    Assert.IsTrue Actual, "Array not sorted"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Sort")
Private Sub Sort_MultiDimArrayQuicksortRecursive_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    TestArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.Items = TestArray
    SUT.SortMethod = SM_QUICKSORT_RECURSIVE
    'Act:
    SUT.Sort
    Actual = SUT.IsSorted()
    'Assert:
    Assert.IsTrue Actual, "Array not sorted"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Sort")
Private Sub Sort_MultiDimArrayQuicksortRecursiveNegativeBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    TestArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.LowerBound = -10
    SUT.Items = TestArray
    SUT.SortMethod = SM_QUICKSORT_RECURSIVE
    'Act:
    SUT.Sort
    Actual = SUT.IsSorted
    
    'Assert:
    Assert.IsTrue Actual, "Array not sorted"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Sort")
Private Sub Sort_MultiDimArrayQuicksortRecursivePositiveBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    TestArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.LowerBound = 10
    SUT.Items = TestArray
    SUT.SortMethod = SM_QUICKSORT_RECURSIVE
    'Act:
    SUT.Sort
    Actual = SUT.IsSorted
    
    'Assert:
    Assert.IsTrue Actual, "Array not sorted"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Sort")
Private Sub Sort_JaggedArrayQuicksortRecursive_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    TestArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.Items = TestArray
    SUT.SortMethod = SM_QUICKSORT_RECURSIVE
    'Act:
    SUT.Sort
    Actual = SUT.IsSorted()
    'Assert:
    Assert.IsTrue Actual, "Array not sorted"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Sort")
Private Sub Sort_JaggedArrayQuicksortRecursiveNegativeBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    TestArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.LowerBound = -10
    SUT.Items = TestArray
    SUT.SortMethod = SM_QUICKSORT_RECURSIVE
    'Act:
    SUT.Sort
    Actual = SUT.IsSorted
    
    'Assert:
    Assert.IsTrue Actual, "Array not sorted"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Sort")
Private Sub Sort_JaggedArrayQuicksortRecursivePositiveBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    TestArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.LowerBound = 10
    SUT.Items = TestArray
    SUT.SortMethod = SM_QUICKSORT_RECURSIVE
    'Act:
    SUT.Sort
    Actual = SUT.IsSorted
    
    'Assert:
    Assert.IsTrue Actual, "Array not sorted"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Sort")
Private Sub Sort_OneDimArrayQuicksortIterative_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    TestArray = Gen.GetArray()
    SUT.Items = TestArray
    SUT.SortMethod = SM_QUICKSORT_ITERATIVE
    'Act:
    SUT.Sort
    Actual = SUT.IsSorted
    
    'Assert:
    Assert.IsTrue Actual, "Array not sorted"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Sort")
Private Sub Sort_OneDimArrayQuicksortIterativeNegativeBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    TestArray = Gen.GetArray()
    SUT.LowerBound = -10
    SUT.Items = TestArray
    SUT.SortMethod = SM_QUICKSORT_ITERATIVE
    'Act:
    SUT.Sort
    Actual = SUT.IsSorted
    
    'Assert:
    Assert.IsTrue Actual, "Array not sorted"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Sort")
Private Sub Sort_OneDimArrayQuicksortIterativePositiveBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    TestArray = Gen.GetArray()
    SUT.LowerBound = 10
    SUT.Items = TestArray
    SUT.SortMethod = SM_QUICKSORT_ITERATIVE
    'Act:
    SUT.Sort
    Actual = SUT.IsSorted
    
    'Assert:
    Assert.IsTrue Actual, "Array not sorted"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Sort")
Private Sub Sort_MultiDimArrayQuicksortIterative_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    TestArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.Items = TestArray
    SUT.SortMethod = SM_QUICKSORT_ITERATIVE
    'Act:
    SUT.Sort
    Actual = SUT.IsSorted()
    'Assert:
    Assert.IsTrue Actual, "Array not sorted"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Sort")
Private Sub Sort_MultiDimArrayQuicksortIterativeNegativeBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    TestArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.LowerBound = -10
    SUT.Items = TestArray
    SUT.SortMethod = SM_QUICKSORT_ITERATIVE
    'Act:
    SUT.Sort
    Actual = SUT.IsSorted
    
    'Assert:
    Assert.IsTrue Actual, "Array not sorted"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Sort")
Private Sub Sort_MultiDimArrayQuicksortIterativePositiveBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    TestArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.LowerBound = 10
    SUT.Items = TestArray
    SUT.SortMethod = SM_QUICKSORT_ITERATIVE
    'Act:
    SUT.Sort
    Actual = SUT.IsSorted
    
    'Assert:
    Assert.IsTrue Actual, "Array not sorted"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Sort")
Private Sub Sort_JaggedArrayQuicksortIterative_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    TestArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.Items = TestArray
    SUT.SortMethod = SM_QUICKSORT_ITERATIVE
    'Act:
    SUT.Sort
    Actual = SUT.IsSorted()
    'Assert:
    Assert.IsTrue Actual, "Array not sorted"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Sort")
Private Sub Sort_JaggedArrayQuicksortIterativeNegativeBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    TestArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.LowerBound = -10
    SUT.Items = TestArray
    SUT.SortMethod = SM_QUICKSORT_ITERATIVE
    'Act:
    SUT.Sort
    Actual = SUT.IsSorted
    
    'Assert:
    Assert.IsTrue Actual, "Array not sorted"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Sort")
Private Sub Sort_JaggedArrayQuicksortIterativePositiveBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    TestArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.LowerBound = 10
    SUT.Items = TestArray
    SUT.SortMethod = SM_QUICKSORT_ITERATIVE
    'Act:
    SUT.Sort
    Actual = SUT.IsSorted
    
    'Assert:
    Assert.IsTrue Actual, "Array not sorted"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Sort")
Private Sub Sort_OneDimArrayTimSort_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    TestArray = Gen.GetArray()
    SUT.Items = TestArray
    SUT.SortMethod = SM_TIMSORT
    'Act:
    SUT.Sort
    Actual = SUT.IsSorted
    
    'Assert:
    Assert.IsTrue Actual, "Array not sorted"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Sort")
Private Sub Sort_OneDimArrayTimSort10kEntries_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    TestArray = Gen.GetArray(Length:=10000)
    SUT.Items = TestArray
    SUT.SortMethod = SM_TIMSORT
    'Act:
    SUT.Sort
    Actual = SUT.IsSorted
    'Assert:
    Assert.IsTrue Actual, "Array not sorted"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Sort")
Private Sub Sort_OneDimArrayTimSortNegativeBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    TestArray = Gen.GetArray()
    SUT.LowerBound = -10
    SUT.Items = TestArray
    SUT.SortMethod = SM_TIMSORT
    'Act:
    SUT.Sort
    Actual = SUT.IsSorted
    
    'Assert:
    Assert.IsTrue Actual, "Array not sorted"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Sort")
Private Sub Sort_OneDimArrayTimSortPositiveBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    TestArray = Gen.GetArray()
    SUT.LowerBound = 10
    SUT.Items = TestArray
    SUT.SortMethod = SM_TIMSORT
    'Act:
    SUT.Sort
    Actual = SUT.IsSorted
    
    'Assert:
    Assert.IsTrue Actual, "Array not sorted"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Sort")
Private Sub Sort_MultiDimArrayTimSort_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    TestArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.Items = TestArray
    SUT.SortMethod = SM_TIMSORT
    'Act:
    SUT.Sort
    Actual = SUT.IsSorted()
    'Assert:
    Assert.IsTrue Actual, "Array not sorted"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Sort")
Private Sub Sort_MultiDimArrayTimSortNegativeBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    TestArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.LowerBound = -10
    SUT.Items = TestArray
    SUT.SortMethod = SM_TIMSORT
    'Act:
    SUT.Sort
    Actual = SUT.IsSorted
    
    'Assert:
    Assert.IsTrue Actual, "Array not sorted"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Sort")
Private Sub Sort_MultiDimArrayTimSortPositiveBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    TestArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.LowerBound = 10
    SUT.Items = TestArray
    SUT.SortMethod = SM_TIMSORT
    'Act:
    SUT.Sort
    Actual = SUT.IsSorted
    
    'Assert:
    Assert.IsTrue Actual, "Array not sorted"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Sort")
Private Sub Sort_JaggedArrayTimSort_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    TestArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.Items = TestArray
    SUT.SortMethod = SM_TIMSORT
    'Act:
    SUT.Sort
    Actual = SUT.IsSorted()
    'Assert:
    Assert.IsTrue Actual, "Array not sorted"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Sort")
Private Sub Sort_JaggedArrayTimSortNegativeBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    TestArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.LowerBound = -10
    SUT.Items = TestArray
    SUT.SortMethod = SM_TIMSORT
    'Act:
    SUT.Sort
    Actual = SUT.IsSorted
    
    'Assert:
    Assert.IsTrue Actual, "Array not sorted"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Sort")
Private Sub Sort_JaggedArrayTimSortPositiveBase_ArrayIsSorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    TestArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.LowerBound = 10
    SUT.Items = TestArray
    SUT.SortMethod = SM_TIMSORT
    'Act:
    SUT.Sort
    Actual = SUT.IsSorted
    
    'Assert:
    Assert.IsTrue Actual, "Array not sorted"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'''''''''''''''''''''''
' Method - CopyWithin '
'''''''''''''''''''''''

'@TestMethod("BetterArray_CopyWithin")
Private Sub CopyWithin_OneDimArrayElement3ToIndex0_SelectionCopiedLengthUnchanged()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    
    TestArray = Array("a", "b", "c", "d", "e")
    Expected = Array("d", "b", "c", "d", "e")
    SUT.Items = TestArray
    'Act:
    Actual = SUT.CopyWithin(0, 3, 4).Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_CopyWithin")
Private Sub CopyWithin_OneDimArrayElements3ToEndToIndex1_SelectionCopiedLengthUnchanged()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    
    TestArray = Array("a", "b", "c", "d", "e")
    Expected = Array("a", "d", "e", "d", "e")
    SUT.Items = TestArray
    'Act:
    Actual = SUT.CopyWithin(1, 3).Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_CopyWithin")
Private Sub CopyWithin_OneDimArrayFirstTwoElementsToLastTwoElements_SelectionCopiedLengthUnchanged()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    
    TestArray = Array("Banana", "Orange", "Apple", "Mango")
    Expected = Array("Banana", "Orange", "Banana", "Orange")
    SUT.Items = TestArray
    'Act:
    Actual = SUT.CopyWithin(2, 0).Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_CopyWithin")
Private Sub CopyWithin_OneDimArrayNoStartNoEnd_NothingChanged()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    
    TestArray = Array("a", "b", "c", "d", "e")
    Expected = Array("a", "b", "c", "d", "e")
    SUT.Items = TestArray
    'Act:
    Actual = SUT.CopyWithin(0).Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_CopyWithin")
Private Sub CopyWithin_OneDimArrayPositiveStartNoEnd_SelectionCopiedLengthUnchanged()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    
    TestArray = Array("a", "b", "c", "d", "e")
    Expected = Array("d", "e", "c", "d", "e")
    SUT.Items = TestArray
    'Act:
    Actual = SUT.CopyWithin(0, 3).Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_CopyWithin")
Private Sub CopyWithin_OneDimArrayNegativeStartNoEnd_SelectionCopiedLengthUnchanged()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    
    TestArray = Array("a", "b", "c", "d", "e")
    Expected = Array("d", "e", "c", "d", "e")
    SUT.Items = TestArray
    'Act:
    Actual = SUT.CopyWithin(0, -2).Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_CopyWithin")
Private Sub CopyWithin_OneDimArrayPositiveStartPositiveEnd_SelectionCopiedLengthUnchanged()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    
    TestArray = Array("a", "b", "c", "d", "e")
    Expected = Array("c", "b", "c", "d", "e")
    SUT.Items = TestArray
    'Act:
    Actual = SUT.CopyWithin(0, 2, 3).Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_CopyWithin")
Private Sub CopyWithin_OneDimArrayPositiveStartNegativeEnd_SelectionCopiedLengthUnchanged()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    
    TestArray = Array("a", "b", "c", "d", "e")
    Expected = Array("c", "d", "c", "d", "e")
    SUT.Items = TestArray
    'Act:
    Actual = SUT.CopyWithin(0, 2, -1).Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_CopyWithin")
Private Sub CopyWithin_OneDimArrayNegativeStartNegativeEnd_SelectionCopiedLengthUnchanged()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    
    TestArray = Array("a", "b", "c", "d", "e")
    Expected = Array("c", "b", "c", "d", "e")
    SUT.Items = TestArray
    'Act:
    Actual = SUT.CopyWithin(0, -3, -2).Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_CopyWithin")
Private Sub CopyWithin_JaggedArrayElement3ToIndex0_SelectionCopiedLengthUnchanged()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestResult As Boolean
    
    TestArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    Expected = TestArray
    Expected(0) = Expected(3)
    SUT.Items = TestArray
    'Act:
    Actual = SUT.CopyWithin(0, 3, 4).Items
    TestResult = SequenceEquals_JaggedArray(Expected, Actual)
    'Assert:
    Assert.IsTrue TestResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_CopyWithin")
Private Sub CopyWithin_EmptyInternal_RaisesError()
    Const ExpectedError As Long = ErrorCodes.EC_UNALLOCATED_ARRAY
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore VariableNotUsed
    Dim Actual() As Variant
    'Act:
    Actual = SUT.CopyWithin(0, 3, 4).Items
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
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
    Dim TestArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    
    TestArray = Array("Foo", "Bar", "Fizz", "Buzz")
    Expected = Array("Foo", "Fizz", "Buzz")
    SUT.Items = TestArray
    'Act:
    Actual = SUT.Filter("Bar").Items

    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Filter")
Private Sub Filter_OneDimInclude_ReturnsFilteredArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    
    TestArray = Array("Foo", "Bar", "Fizz", "Buzz")
    Expected = Array("Bar")
    SUT.Items = TestArray
    'Act:
    Actual = SUT.Filter("Bar", True).Items

    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Filter")
Private Sub Filter_JaggedArrayExclude_ReturnsFilteredArray()
    On Error GoTo TestFail

    'Arrange:
    Dim TestArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestResult As Boolean

    TestArray = Array(Array("Foo", "Bar"), Array("Fizz", "Buzz"))
    Expected = Array(Array("Foo"), Array("Fizz", "Buzz"))
    SUT.Items = TestArray
    'Act:
    Actual = SUT.Filter("Bar", False).Items
    TestResult = SequenceEquals_JaggedArray(Expected, Actual)
    'Assert:
    Assert.IsTrue TestResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Filter")
Private Sub Filter_JaggedArrayInclude_ReturnsFilteredArrayn()
    On Error GoTo TestFail

    'Arrange:
    Dim TestArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestResult As Boolean

    TestArray = Array(Array("Foo", "Bar"), Array("Fizz", "Buzz"))
    Expected = Array(Array("Bar"))

    SUT.Items = TestArray
    'Act:
    Actual = SUT.Filter("Bar", True, True).Items
    TestResult = SequenceEquals_JaggedArray(Expected, Actual)
    'Assert:
    Assert.IsTrue TestResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Filter")
Private Sub Filter_MultiDimArrayExclude_ReturnsFilteredArray()
    On Error GoTo TestFail

    'Arrange:
    Dim TestArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant

    ReDim TestArray(1 To 2, 1 To 2)
    TestArray(1, 1) = "Foo"
    TestArray(1, 2) = "Bar"
    TestArray(2, 1) = "Fizz"
    TestArray(2, 2) = "Buzz"

    ReDim Expected(1 To 2, 1 To 2)
    Expected(1, 1) = "Foo"
    Expected(2, 1) = "Fizz"
    Expected(2, 2) = "Buzz"

    SUT.Items = TestArray
    'Act:
    SUT.Filter "Bar", False, True
    Actual = SUT.Items

    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Filter")
Private Sub Filter_MultiDimArrayInclude_ReturnsFilteredArray()
    On Error GoTo TestFail

    'Arrange:
    Dim TestArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant

    ReDim TestArray(1 To 2, 1 To 2)
    TestArray(1, 1) = "Foo"
    TestArray(1, 2) = "Bar"
    TestArray(2, 1) = "Fizz"
    TestArray(2, 2) = "Buzz"

    ReDim Expected(1 To 1, 1 To 1)
    Expected(1, 1) = "Bar"

    SUT.Items = TestArray
    'Act:
    Actual = SUT.Filter("Bar", True, True).Items
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'''''''''''''''''''''''
' Method - FilterType '
'''''''''''''''''''''''

'@TestMethod("BetterArray_FilterType")
Private Sub FilterType_OneDimExclude_ReturnsFilteredArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    
    TestArray = Array("Foo", 1.23, "Fizz", "Buzz")
    Expected = Array("Foo", "Fizz", "Buzz")
    SUT.Items = TestArray
    'Act:
    Actual = SUT.FilterType("double").Items

    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_FilterType")
Private Sub FilterType_OneDimInclude_ReturnsFilteredArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    
    TestArray = Array(1, "Bar", 1.23, 100)
    Expected = Array("Bar")
    SUT.Items = TestArray
    'Act:
    Actual = SUT.FilterType("string", True).Items

    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_FilterType")
Private Sub FilterType_JaggedArrayExclude_ReturnsFilteredArray()
    On Error GoTo TestFail

    'Arrange:
    Dim TestArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestResult As Boolean

    TestArray = Array(Array("Foo", 1.5), Array("Fizz", "Buzz"))
    Expected = Array(Array("Foo"), Array("Fizz", "Buzz"))
    SUT.Items = TestArray
    'Act:
    Actual = SUT.FilterType("double", False).Items
    TestResult = SequenceEquals_JaggedArray(Expected, Actual)
    'Assert:
    Assert.IsTrue TestResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_FilterType")
Private Sub FilterType_JaggedArrayInclude_ReturnsFilteredArray()
    On Error GoTo TestFail

    'Arrange:
    Dim TestArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestResult As Boolean

    TestArray = Array(Array(1, "Bar"), Array(1.2, -4))
    Expected = Array(Array("Bar"))

    SUT.Items = TestArray
    'Act:
    Actual = SUT.FilterType("string", True, True).Items
    TestResult = SequenceEquals_JaggedArray(Expected, Actual)
    'Assert:
    Assert.IsTrue TestResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_FilterType")
Private Sub FilterType_MultiDimArrayExclude_ReturnsFilteredArray()
    On Error GoTo TestFail

    'Arrange:
    Dim TestArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant

    ReDim TestArray(1 To 2, 1 To 2)
    TestArray(1, 1) = "Foo"
    TestArray(1, 2) = 1.23
    TestArray(2, 1) = "Fizz"
    TestArray(2, 2) = "Buzz"

    ReDim Expected(1 To 2, 1 To 2)
    Expected(1, 1) = "Foo"
    Expected(2, 1) = "Fizz"
    Expected(2, 2) = "Buzz"

    SUT.Items = TestArray
    'Act:
    SUT.FilterType "double", False, True
    Actual = SUT.Items

    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_FilterType")
Private Sub FilterType_MultiDimArrayInclude_ReturnsFilteredArray()
    On Error GoTo TestFail

    'Arrange:
    Dim TestArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant

    ReDim TestArray(1 To 2, 1 To 2)
    TestArray(1, 1) = 1.23
    TestArray(1, 2) = "Bar"
    TestArray(2, 1) = 123
    TestArray(2, 2) = 5000

    ReDim Expected(1 To 1, 1 To 1)
    Expected(1, 1) = "Bar"

    SUT.Items = TestArray
    'Act:
    Actual = SUT.FilterType("string", True, True).Items
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'''''''''''''''''''''
' Method - Includes '
'''''''''''''''''''''

'@TestMethod("BetterArray_Includes")
Private Sub Includes_OneDimArrayContainsTarget_ReturnsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim Expected As Boolean
    Dim Actual As Boolean
    TestArray = Array("Foo", "Bar", "Fizz", "Buzz")
    Expected = True
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.Includes("Bar")
    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Includes")
Private Sub Includes_OneDimArrayDoesntContainTarget_ReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim Expected As Boolean
    Dim Actual As Boolean
    TestArray = Array("Foo", "Bar", "Fizz", "Buzz")
    Expected = False
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.Includes("wibble")
    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Includes")
Private Sub Includes_OneDimArrayDoesContainTargetAfterStartIndex_ReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim Expected As Boolean
    Dim Actual As Boolean
    TestArray = Array("Foo", "Bar", "Fizz", "Buzz")
    Expected = True
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.Includes("Fizz", 2)
    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Includes")
Private Sub Includes_OneDimArrayDoesntContainTargetAfterStartIndex_ReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim Expected As Boolean
    Dim Actual As Boolean
    TestArray = Array("Foo", "Bar", "Fizz", "Buzz")
    Expected = False
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.Includes("Foo", 2)
    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Includes")
Private Sub Includes_JaggedArrayContains_ReturnsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim Expected As Boolean
    Dim Actual As Boolean
    TestArray = Array(Array("Foo", "Bar"), Array("Fizz", "Buzz"))
    Expected = True
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.Includes("Buzz", Recurse:=True)
    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Includes")
Private Sub Includes_JaggedArrayDoesntContains_ReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim Expected As Boolean
    Dim Actual As Boolean
    TestArray = Array(Array("Foo", "Bar"), Array("Fizz", "Buzz"))
    Expected = False
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.Includes("wibble", Recurse:=True)
    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Includes")
Private Sub Includes_EmptyInternal_ReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Boolean
    Dim Actual As Boolean
    Expected = False
    
    'Act:
    Actual = SUT.Includes("Foo")
    
    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'''''''''''''''''''''''''
' Method - IncludesType '
'''''''''''''''''''''''''

'@TestMethod("BetterArray_IncludesType")
Private Sub IncludesType_OneDimArrayContainsType_ReturnsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim Expected As Boolean
    Dim Actual As Boolean
    Dim SearchType As String
    TestArray = Gen.GetArray(AG_DOUBLE)
    Expected = True
    SUT.Items = TestArray
    SearchType = "Double"
    
    'Act:
    Actual = SUT.IncludesType(SearchType)
    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_IncludesType")
Private Sub IncludesType_OneDimArrayDoesntContainType_ReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim Expected As Boolean
    Dim Actual As Boolean
    Dim SearchType As String
    TestArray = Gen.GetArray(AG_DOUBLE)
    Expected = False
    SUT.Items = TestArray
    SearchType = "string"
    
    'Act:
    Actual = SUT.IncludesType(SearchType)
    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_IncludesType")
Private Sub IncludesType_JaggedArrayContainsType_ReturnsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim Expected As Boolean
    Dim Actual As Boolean
    Dim SearchType As String
    TestArray = Gen.GetArray(AG_DOUBLE)
    Expected = True
    SUT.Items = TestArray
    SearchType = "Double"
    
    'Act:
    Actual = SUT.IncludesType(SearchType, Recurse:=True)
    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'''''''''''''''''
' Method - Keys '
'''''''''''''''''

'@TestMethod("BetterArray_Keys")
Private Sub Keys_OneDimArrayDefaultBase_ReturnsCorrectKeys()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim i As Long
    
    TestArray = Gen.GetArray
    ReDim Expected(LBound(TestArray) To UBound(TestArray))
    For i = LBound(TestArray) To UBound(TestArray)
        Expected(i) = i
    Next
    SUT.Items = TestArray
    'Act:
    Actual = SUT.Keys
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Keys")
Private Sub Keys_OneDimArraySpecifiedBase_ReturnsCorrectKeys()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim i As Long
    
    SUT.LowerBound = 2
    TestArray = Gen.GetArray
    ReDim Expected(0 To Gen.GetArrayLength(TestArray) - 1)
    For i = LBound(Expected) To UBound(Expected)
        Expected(i) = i + 2
    Next
    SUT.Items = TestArray
    'Act:
    Actual = SUT.Keys
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Keys")
Private Sub Keys_MultiDimArray_ReturnsCorrectKeys()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim i As Long
    
    TestArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    ReDim Expected(LBound(TestArray) To UBound(TestArray))
    For i = LBound(TestArray) To UBound(TestArray)
        Expected(i) = i
    Next
    SUT.Items = TestArray
    'Act:
    Actual = SUT.Keys
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Keys")
Private Sub Keys_JaggedArray_ReturnsCorrectKeys()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim i As Long
    
    TestArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    ReDim Expected(LBound(TestArray) To UBound(TestArray))
    For i = LBound(TestArray) To UBound(TestArray)
        Expected(i) = i
    Next
    SUT.Items = TestArray
    'Act:
    Actual = SUT.Keys
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Keys")
Private Sub Keys_EmptyInternal_RaisesUnallocError()
    Const ExpectedError As Long = ErrorCodes.EC_UNALLOCATED_ARRAY
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore VariableNotUsed
    Dim Actual() As Variant
    
    'Act:
    Actual = SUT.Keys
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
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
    Dim TestArray() As Variant
    TestArray = Array(1, 3, 2, 6, 4, 9, 0, 5)
    Dim Expected As Long
    Dim Actual As Long
    
    Expected = 9
    SUT.Items = TestArray
    'Act:
    Actual = SUT.Max

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Max")
Private Sub Max_OneDimArrayStringsInternal_ReturnsLargest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    TestArray = Array("Foo", "Bar", "Fizz", "Buzz")
    Dim Expected As String
    Dim Actual As String
    
    Expected = "Foo"
    SUT.Items = TestArray
    'Act:
    Actual = CStr(SUT.Max)

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Max")
Private Sub Max_OneDimArrayVariantsInternal_ReturnsLargest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    TestArray = Array("Foo", 1, "Bar", 100, "Fizz", -1, "Buzz")
    Dim Expected As Variant
    Dim Actual As Variant
    Dim TestResult As Boolean
    
    Expected = "Foo"
    SUT.Items = TestArray
    'Act:
    Actual = SUT.Max
    TestResult = ElementsAreEqual(Expected, Actual)
    'Assert:
    Assert.IsTrue TestResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Max")
Private Sub Max_OneDimArrayObjects_ReturnsEmpty()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    TestArray = Gen.GetArray(AG_OBJECT)
    Dim Expected As Variant
    Dim Actual As Variant
    
    Expected = Empty
    SUT.Items = TestArray
    'Act:
    Actual = SUT.Max

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Max")
Private Sub Max_ParamArray_ReturnsLargest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Variant
    Dim Actual As Variant
    Dim TestResult As Boolean
    
    Expected = "Foo"
    'Act:
    Actual = SUT.Max("Foo", 1, "Bar", 100, "Fizz", -1, "Buzz")
    TestResult = ElementsAreEqual(Expected, Actual)
    
    'Assert:
    Assert.IsTrue TestResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Max")
Private Sub Max_PassedArray_ReturnsLargest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    TestArray = Array("Foo", 1, "Bar", 100, "Fizz", -1, "Buzz")
    Dim Expected As Variant
    Dim Actual As Variant
    Dim TestResult As Boolean
    
    Expected = "Foo"
    'Act:
    Actual = SUT.Max(TestArray)
    TestResult = ElementsAreEqual(Expected, Actual)
    
    'Assert:
    Assert.IsTrue TestResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Max")
Private Sub Max_JaggedArray_Returnslargest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Long
    Dim Actual As Long
    Dim TestArray As Variant
    TestArray = Array(Array(1, 3, 20, 4), Array(8, 2, 7, 9))
    Expected = 20
    'Act:
    Actual = SUT.Max(TestArray)

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Max")
Private Sub Max_EmptyInternal_ReturnsEmpty()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Variant
    Dim Expected As Variant
    Expected = Empty
    
    'Act:
    Actual = SUT.Max

    'Assert:
    Assert.AreSame Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


''''''''''''''''
' Method - Min '
''''''''''''''''

'@TestMethod("BetterArray_Min")
Private Sub Min_OneDimArrayNumericInternal_ReturnsSmallest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    TestArray = Array(1, 3, 2, 6, 4, 9, 0, 5)
    Dim Expected As Long
    Dim Actual As Long
    
    Expected = 0
    SUT.Items = TestArray
    'Act:
    Actual = SUT.Min

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Min")
Private Sub Min_OneDimArrayStringsInternal_ReturnsSmallest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    TestArray = Array("Foo", "Bar", "Fizz", "Buzz")
    Dim Expected As String
    Dim Actual As String
    
    Expected = "Bar"
    SUT.Items = TestArray
    'Act:
    Actual = CStr(SUT.Min)

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Min")
Private Sub Min_OneDimArrayVariantsInternal_ReturnsSmallest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    TestArray = Array("Foo", 1, "Bar", 100, "Fizz", -1, "Buzz")
    Dim Expected As Variant
    Dim Actual As Variant
    Dim TestResult As Boolean
    
    Expected = -1
    SUT.Items = TestArray
    'Act:
    Actual = SUT.Min
    TestResult = ElementsAreEqual(Expected, Actual)
    'Assert:
    Assert.IsTrue TestResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Min")
Private Sub Min_OneDimArrayObjects_ReturnsEmpty()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    TestArray = Gen.GetArray(AG_OBJECT)
    Dim Expected As Variant
    Dim Actual As Variant
    
    Expected = Empty
    SUT.Items = TestArray
    'Act:
    Actual = SUT.Min

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Min")
Private Sub Min_ParamArray_ReturnsSmallest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Variant
    Dim Actual As Variant
    Dim TestResult As Boolean
    
    Expected = -1
    'Act:
    Actual = SUT.Min("Foo", 1, "Bar", 100, "Fizz", -1, "Buzz")
    TestResult = ElementsAreEqual(Expected, Actual)
    
    'Assert:
    Assert.IsTrue TestResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Min")
Private Sub Min_PassedArray_ReturnsSmallest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    TestArray = Array("Foo", 1, "Bar", 100, "Fizz", -1, "Buzz")
    Dim Expected As Variant
    Dim Actual As Variant
    Dim TestResult As Boolean
    
    Expected = -1
    'Act:
    Actual = SUT.Min(TestArray)
    TestResult = ElementsAreEqual(Expected, Actual)
    
    'Assert:
    Assert.IsTrue TestResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Min")
Private Sub Min_JaggedArray_ReturnsSmallest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Long
    Dim Actual As Long
    Dim TestArray As Variant
    TestArray = Array(Array(1, 3, 20, 4), Array(8, 2, 7, 9))
    Expected = 1
    'Act:
    Actual = SUT.Min(TestArray)

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Min")
Private Sub Min_EmptyInternal_ReturnsEmpty()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Variant
    Dim Expected As Variant
    Expected = Empty
    
    'Act:
    Actual = SUT.Min

    'Assert:
    Assert.AreSame Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

''''''''''''''''''
' Method - Slice '
''''''''''''''''''

'@TestMethod("BetterArray_Slice")
Private Sub Slice_OneDimNoEndArg_ReturnsCopy()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    TestArray = Gen.GetArray(AG_VARIANT)
    Expected = TestArray
    
    SUT.Items = TestArray
    'Act:
    Actual = SUT.Slice(0)
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Slice")
Private Sub Slice_OneDimNoEndArgObjects_ReturnsCopy()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    TestArray = Gen.GetArray(AG_OBJECT)
    Expected = TestArray
    
    SUT.Items = TestArray
    'Act:
    Actual = SUT.Slice(0)
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Slice")
Private Sub Slice_OneDimWithEndArg_ReturnsCopy()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    TestArray = Array("Foo", "Bar", "Fizz", "Buzz")
    Expected = Array("Foo", "Bar")
    
    SUT.Items = TestArray
    'Act:
    Actual = SUT.Slice(0, 2)
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Slice")
Private Sub Slice_MultiDimNoEndArg_ReturnsCopy()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray(1 To 4, 1 To 2) As Variant
    TestArray(1, 1) = "Foo"
    TestArray(1, 2) = "Bar"
    TestArray(2, 1) = "Fizz"
    TestArray(2, 2) = "Buzz"
    TestArray(3, 1) = "Xyzzy"
    TestArray(3, 2) = "flob"
    TestArray(4, 1) = "quux"
    TestArray(4, 2) = "quuz"
    
    Expected = TestArray
    
    SUT.Items = TestArray
    'Act:
    Actual = SUT.Slice(1)
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Slice")
Private Sub Slice_MultiDimWithEndArg_ReturnsCopy()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected(1 To 2, 1 To 2) As Variant
    Dim Actual() As Variant
    Dim TestArray(1 To 4, 1 To 2) As Variant
    TestArray(1, 1) = "Foo"
    TestArray(1, 2) = "Bar"
    TestArray(2, 1) = "Fizz"
    TestArray(2, 2) = "Buzz"
    TestArray(3, 1) = "Xyzzy"
    TestArray(3, 2) = "flob"
    TestArray(4, 1) = "quux"
    TestArray(4, 2) = "quuz"
    
    Expected(1, 1) = "Foo"
    Expected(1, 2) = "Bar"
    Expected(2, 1) = "Fizz"
    Expected(2, 2) = "Buzz"
    
    SUT.Items = TestArray
    'Act:
    Actual = SUT.Slice(1, 3)
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Slice")
Private Sub Slice_JaggedNoEndArg_ReturnsCopy()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestResult As Boolean
    Expected = Gen.GetArray(ArrayType:=AG_JAGGED)

    SUT.Items = Expected
    'Act:
    Actual = SUT.Slice(LBound(Expected))
    TestResult = SequenceEquals_JaggedArray(Expected, Actual)
    'Assert:
    Assert.IsTrue TestResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Slice")
Private Sub Slice_JaggedWithEndArg_ReturnsCopy()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim TestResult As Boolean
    
    TestArray = Array(Array("Foo", "Bar"), Array("Fizz", "Buzz"), _
        Array("Xyzzy", "flob"), Array("quux", "quuz"))
   
    Expected = Array(Array("Foo", "Bar"), Array("Fizz", "Buzz"))
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.Slice(LBound(Expected), 2)
    TestResult = SequenceEquals_JaggedArray(Expected, Actual)
    
    'Assert:
    Assert.IsTrue TestResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Slice")
Private Sub Slice_EmptyInternal_GracefulDegradation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant

    'Act:
    Actual = SUT.Slice(1)
    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

''''''''''''''''''''
' Method - Reverse '
''''''''''''''''''''

'@TestMethod("BetterArray_Reverse")
Private Sub Reverse_OneDimArray_ArrayIsReversed()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestResult As Boolean
    
    Expected = Gen.GetArray
    SUT.Items = Expected
    
    'Act:
    Actual = SUT.Reverse.Items
    TestResult = arraysAreReversed(Expected, Actual)
    'Assert:
    Assert.IsTrue TestResult, "Actual not reverse of expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Reverse")
Private Sub Reverse_OneDimArrayBase10_ArrayIsReversed()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestResult As Boolean
    
    Gen.LowerBound = 10
    Expected = Gen.GetArray
    SUT.Items = Expected
    'Act:
    Actual = SUT.Reverse.Items
    TestResult = arraysAreReversed(Expected, Actual)
    'Assert:
    Assert.IsTrue TestResult, "Actual not reverse of expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Reverse")
Private Sub Reverse_MultiDimArray_ArrayIsReversed()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestResult As Boolean
    Dim i As Long

    Expected = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.Items = Expected
    'Act:
    Actual = SUT.Reverse.Items
    TestResult = True
    For i = LBound(Expected) To UBound(Expected)
        If Not ElementsAreEqual( _
                Expected(i, LBound(Expected, 2)), _
                Actual(LBound(Expected) + UBound(Expected) - i, LBound(Expected, 2)) _
            ) Then
            TestResult = False
            Exit For
        End If
    Next
    
    'Assert:
    Assert.IsTrue TestResult, "Actual not reverse of expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Reverse")
Private Sub Reverse_MultiDimArrayRecursive_ArrayIsReversed()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestResult As Boolean
    Dim i As Long
    
    Expected = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.Items = Expected
    'Act:
    Actual = SUT.Reverse(True).Items
    TestResult = True
    For i = LBound(Expected) To UBound(Expected)
        If Not ElementsAreEqual( _
                Expected(i, LBound(Expected, 2)), _
                Actual(LBound(Expected) + UBound(Expected) - i, UBound(Expected, 2)) _
            ) Then
            TestResult = False
            Exit For
        End If
    Next
    
    'Assert:
    Assert.IsTrue TestResult, "Actual not reverse of expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Reverse")
Private Sub Reverse_JaggedArray_ArrayIsReversed()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestResult As Boolean
    
    Expected = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.Items = Expected
    'Act:
    Actual = SUT.Reverse.Items
    TestResult = arraysAreReversed(Expected, Actual)
    'Assert:
    Assert.IsTrue TestResult, "Actual not reverse of expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Reverse")
Private Sub Reverse_JaggedArrayRecurse_ArrayIsReversed()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestResult As Boolean
    
    Expected = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.Items = Expected
    'Act:
    Actual = SUT.Reverse(True).Items
    TestResult = arraysAreReversed(Expected, Actual, True)
    'Assert:
    Assert.IsTrue TestResult, "Actual not reverse of expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Reverse")
Private Sub Reverse_EmptyInternal_ReturnsEmpty()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    ReDim Expected(0) As Variant
    Expected(0) = Empty
    'Act:
    Actual = SUT.Reverse.Items

    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub



''''''''''''''''''''
' Method - Shuffle '
''''''''''''''''''''

'@TestMethod("BetterArray_Shuffle")
Private Sub Shuffle_OneDimArray_ArrayIsShuffled()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim SortedArray() As Variant
    Dim Actual() As Variant
    
    TestArray = Gen.GetArray(AG_DOUBLE)
    SUT.Items = TestArray
    SortedArray = SUT.Sort.Items
    'Act:
    Actual = SUT.Shuffle.Items

    'Assert:
    Assert.NotSequenceEquals SortedArray, Actual, "Array is not shufled"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Shuffle")
Private Sub Shuffle_OneDimArrayBase1_ArrayIsShuffled()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim SortedArray() As Variant
    Dim Actual() As Variant
    
    Gen.LowerBound = 1
    TestArray = Gen.GetArray(AG_DOUBLE)
    SUT.Items = TestArray
    SortedArray = SUT.Sort.Items
    'Act:
    Actual = SUT.Shuffle.Items

    'Assert:
    Assert.NotSequenceEquals SortedArray, Actual, "Array is not shufled"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Shuffle")
Private Sub Shuffle_MultiDimArray_ArrayIsShuffled()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim SortedArray() As Variant
    Dim Actual() As Variant
    
    TestArray = Gen.GetArray(AG_DOUBLE, AG_MULTIDIMENSION)
    SUT.Items = TestArray
    SortedArray = SUT.Sort.Items
    'Act:
    Actual = SUT.Shuffle.Items

    'Assert:
    Assert.NotSequenceEquals SortedArray, Actual, "Array is not shufled"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Shuffle")
Private Sub Shuffle_JaggedArray_ArrayIsShuffled()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestArray() As Variant
    Dim SortedArray() As Variant
    Dim Actual() As Variant
    Dim TestResult As Boolean
    
    TestArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.Items = TestArray
    SortedArray = SUT.Sort.Items
    'Act:
    Actual = SUT.Shuffle.Items
    TestResult = SequenceEquals_JaggedArray(SortedArray, Actual)
    'Assert:
    Assert.IsFalse TestResult, "Array is not shufled"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Shuffle")
Private Sub Shuffle_EmptyInternal_ReturnsEmptyArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    
    'Act:
    ReDim Expected(0)
    Actual = SUT.Shuffle.Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'''''''''''''''''''''''''''
' Method - FromExcelRange '
'''''''''''''''''''''''''''

'@TestMethod("BetterArray_FromExcelRange")
Private Sub FromExcelRange_NoDetection_Success()
    On Error GoTo TestFail
    
    'Arrange:
    Dim MockData() As Variant
    Gen.LowerBound = 1
    ' Using an array of strings as Excel converts all numbers to double
    MockData = Gen.GetArray(AG_STRING, AG_MULTIDIMENSION)
    
    Dim LastRow As Long
    LastRow = UBound(MockData, 1)
    Dim LastColumn As Long
    LastColumn = UBound(MockData, 2)
    
    Dim ExcelApp As ExcelProvider
    Set ExcelApp = New ExcelProvider
    ExcelApp.CurrentWorksheet.Range("A1").Resize(LastRow, LastColumn).Value = MockData
    
    Dim Expected(1 To 2, 1 To 2) As Variant
    Expected(1, 1) = MockData(1, 1)
    Expected(1, 2) = MockData(1, 2)
    Expected(2, 1) = MockData(2, 1)
    Expected(2, 2) = MockData(2, 2)
    
    Dim Actual() As Variant
    
    'Act:
    SUT.FromExcelRange ExcelApp.CurrentWorksheet.Range("A1:B2")
    Actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
    
TestExit:
    Exit Sub
TestFail:
    Debug.Print EXCEL_DEPENDENCY_WARNING
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_FromExcelRange")
Private Sub FromExcelRange_ColumnDetection_Success()
    On Error GoTo TestFail
    
    'Arrange:
    Dim MockData() As Variant
    Gen.LowerBound = 1
    ' Using an array of strings as Excel converts all numbers to double
    MockData = Gen.GetArray(AG_STRING, AG_MULTIDIMENSION)
    
    Dim LastRow As Long
    LastRow = UBound(MockData, 1)
    Dim LastColumn As Long
    LastColumn = UBound(MockData, 2)
    
    Dim ExcelApp As ExcelProvider
    Set ExcelApp = New ExcelProvider
    ExcelApp.CurrentWorksheet.Range("A1").Resize(LastRow, LastColumn).Value = MockData
    
    Dim i As Long
    Dim Expected() As Variant
    ReDim Expected(1 To LastRow)
    For i = 1 To LastRow
        Expected(i) = MockData(1, i)
    Next
    
    Dim Actual() As Variant
    
    'Act:
    SUT.FromExcelRange ExcelApp.CurrentWorksheet.Range("A1"), False, True
    Actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
    
TestExit:
    Exit Sub
TestFail:
    Debug.Print EXCEL_DEPENDENCY_WARNING
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_FromExcelRange")
Private Sub FromExcelRange_RowDetection_Success()
    On Error GoTo TestFail
    
    'Arrange:
    Dim MockData() As Variant
    Gen.LowerBound = 1
    ' Using an array of strings as Excel converts all numbers to double
    MockData = Gen.GetArray(AG_STRING, AG_MULTIDIMENSION)
    
    Dim LastRow As Long
    LastRow = UBound(MockData, 1)
    Dim LastColumn As Long
    LastColumn = UBound(MockData, 2)
    
    Dim ExcelApp As ExcelProvider
    Set ExcelApp = New ExcelProvider
    ExcelApp.CurrentWorksheet.Range("A1").Resize(LastRow, LastColumn).Value = MockData
    
    Dim i As Long
    Dim Expected() As Variant
    ReDim Expected(1 To LastRow)
    For i = 1 To LastRow
        Expected(i) = MockData(i, 1)
    Next
    
    Dim Actual() As Variant
    
    'Act:
    SUT.FromExcelRange ExcelApp.CurrentWorksheet.Range("A1"), True, False
    Actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
    
TestExit:
    Exit Sub
TestFail:
    Debug.Print EXCEL_DEPENDENCY_WARNING
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_FromExcelRange")
Private Sub FromExcelRange_ColumnAndRowDetection_Success()
    On Error GoTo TestFail
    
    'Arrange:
    Dim MockData() As Variant
    Gen.LowerBound = 1
    ' Using an array of strings as Excel converts all numbers to double
    MockData = Gen.GetArray(AG_STRING, AG_MULTIDIMENSION)
    
    Dim LastRow As Long
    LastRow = UBound(MockData, 1)
    Dim LastColumn As Long
    LastColumn = UBound(MockData, 2)
    
    Dim ExcelApp As ExcelProvider
    Set ExcelApp = New ExcelProvider
    ExcelApp.CurrentWorksheet.Range("A1").Resize(LastRow, LastColumn).Value = MockData
    
    Dim Expected() As Variant
    Expected = MockData
    
    Dim Actual() As Variant
    
    'Act:
    SUT.FromExcelRange ExcelApp.CurrentWorksheet.Range("A1"), True, True
    Actual = SUT.Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
    
TestExit:
    Exit Sub
TestFail:
    Debug.Print EXCEL_DEPENDENCY_WARNING
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'''''''''''''''''''''''''
' Method - ToExcelRange '
'''''''''''''''''''''''''

'@TestMethod("BetterArray_ToExcelRange")
Private Sub ToExcelRange_OneDimensionNotTransposed_WritesValuesCorrectly()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Destination As Object
    Dim ReturnedRange As Object
    Dim ExcelApp As ExcelProvider
    Dim Expected() As Variant
    Dim Actual(TEST_ARRAY_LENGTH - 1) As Variant
    
    Set ExcelApp = New ExcelProvider
    Set Destination = ExcelApp.CurrentWorksheet.Range("A1")
    
    ' Use Array of Doubles as all values returned from an Excel range are of type Double
    Expected = Gen.GetArray(AG_DOUBLE)
    SUT.Items = Expected
    
    'Act:
    Set ReturnedRange = SUT.ToExcelRange(Destination)
    Dim i As Long
    For i = 1 To ReturnedRange.Rows.Count
        Actual(i - 1) = ReturnedRange.Cells.Item(i, 1).Value
    Next
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Debug.Print EXCEL_DEPENDENCY_WARNING
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_ToExcelRange")
Private Sub ToExcelRange_OneDimensionTransposed_WritesValuesCorrectly()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Destination As Object
    Dim ReturnedRange As Object
    Dim ExcelApp As ExcelProvider
    Dim Expected() As Variant
    Dim Actual(TEST_ARRAY_LENGTH - 1) As Variant

    Set ExcelApp = New ExcelProvider
    Set Destination = ExcelApp.CurrentWorksheet.Range("A1")
    
    ' Use Array of Doubles as all values returned from an Excel range are of type Double
    Expected = Gen.GetArray(AG_DOUBLE)
    SUT.Items = Expected
    
    'Act:
    Set ReturnedRange = SUT.ToExcelRange(Destination, True)
    Dim i As Long
    For i = 1 To ReturnedRange.Columns.Count
        Actual(i - 1) = ReturnedRange.Cells.Item(1, i).Value
    Next
    

    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Debug.Print EXCEL_DEPENDENCY_WARNING
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_ToExcelRange")
Private Sub ToExcelRange_TwoDimensionNotTransposed_WritesValuesCorrectly()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Destination As Object
    Dim ExcelApp As ExcelProvider
    Dim Expected() As Variant
    Dim Actual As Object
    Dim TestResult As Boolean
    Dim Transposed As Boolean

    Set ExcelApp = New ExcelProvider
    Set Destination = ExcelApp.CurrentWorksheet.Range("A1")
    
    Expected = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.Items = Expected
    Transposed = False
    
    'Act:
    Set Actual = SUT.ToExcelRange(Destination, Transposed)
    TestResult = SequenceEquals_JaggedArrayVsRange(Expected, Actual, Transposed)
    
    'Assert:
    Assert.IsTrue TestResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Debug.Print EXCEL_DEPENDENCY_WARNING
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_ToExcelRange")
Private Sub ToExcelRange_TwoDimensionTransposed_WritesValuesCorrectly()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Destination As Object
    Dim ExcelApp As ExcelProvider
    Dim Expected() As Variant
    Dim Actual As Object
    Dim TestResult As Boolean
    Dim Transposed As Boolean

    Set ExcelApp = New ExcelProvider
    Set Destination = ExcelApp.CurrentWorksheet.Range("A1")
    
    Expected = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.Items = Expected
    Transposed = True
    
    'Act:
    Set Actual = SUT.ToExcelRange(Destination, Transposed)
    TestResult = SequenceEquals_JaggedArrayVsRange(Expected, Actual, Transposed)
    
    'Assert:
    Assert.IsTrue TestResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Debug.Print EXCEL_DEPENDENCY_WARNING
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_ToExcelRange")
Private Sub ToExcelRange_JaggedDepthOfThree_WritesScalarRepresentationOfThirdDimension()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TempBetterArray As BetterArray
    Dim Destination As Object
    Dim ReturnedRange As Object
    Dim OutputSheet As Object
    Dim ExcelApp As ExcelProvider
    Dim i As Long
    Dim j As Long
    Dim Expected(TEST_ARRAY_LENGTH - 1, TEST_ARRAY_LENGTH - 1) As Variant
    Dim Actual(TEST_ARRAY_LENGTH - 1, TEST_ARRAY_LENGTH - 1) As Variant
    Dim SourceArray() As Variant
    
    Set ExcelApp = New ExcelProvider
    Set OutputSheet = ExcelApp.CurrentWorksheet
    Set Destination = OutputSheet.Range("A1")
    
    ' Use Array of Doubles as all values returned from an Excel range are of type Double
    SourceArray = Gen.GetArray(AG_DOUBLE, AG_JAGGED, Depth:=3)
    
    For i = LBound(SourceArray) To UBound(SourceArray)
        For j = LBound(SourceArray(i)) To UBound(SourceArray(i))
            Set TempBetterArray = New BetterArray
            TempBetterArray.Items = SourceArray(i)(j)
            Expected(i, j) = TempBetterArray.ToString()
            Set TempBetterArray = Nothing
        Next
    Next
    
    SUT.Items = SourceArray
    
    'Act:
    Set ReturnedRange = SUT.ToExcelRange(Destination)
    
    For i = 1 To ReturnedRange.Rows.Count
        For j = 1 To ReturnedRange.Columns.Count
            Actual(i - 1, j - 1) = ReturnedRange.Cells.Item(i, j).Value
        Next
    Next

    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Debug.Print EXCEL_DEPENDENCY_WARNING
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

''''''''''''''''''''''''''''
' Method - ParseFromString '
''''''''''''''''''''''''''''

'@TestMethod("BetterArray_ParseFromString")
Private Sub ParseFromString_OneDimensionArrayFromToString_ReturnsCorrectValues()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TempBetterArray As BetterArray
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim SourceString As String
    Dim TestResult As Boolean

    Set TempBetterArray = New BetterArray
    Expected = Gen.GetArray()
    TempBetterArray.Items = Expected
    SourceString = TempBetterArray.ToString()
    
    'Act:
    Actual = SUT.ParseFromString(SourceString).Items
    
    ' can't use Assert.SequenceEquals due to type comparison - Bytes Will be Long in actual
    TestResult = SequenceEquals_JaggedArray(Expected, Actual)
    
    'Assert:
    Assert.IsTrue TestResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_ParseFromString")
Private Sub ParseFromString_Jagged2DeepArrayFromToString_ReturnsCorrectValues()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TempBetterArray As BetterArray
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim SourceString As String
    Dim TestResult As Boolean
    
    Set TempBetterArray = New BetterArray
    
    Expected = Gen.GetArray(AG_BYTE, AG_JAGGED)
    TempBetterArray.Items = Expected
    SourceString = TempBetterArray.ToString()
    
    'Act:
    Actual = SUT.ParseFromString(SourceString).Items
    
    ' can't use Assert.SequenceEquals due to type comparison - Bytes Will be Long in actual
    TestResult = SequenceEquals_JaggedArray(Expected, Actual)
    
    'Assert:
    Assert.IsTrue TestResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_ParseFromString")
Private Sub ParseFromString_Jagged3DeepArrayFromToString_ReturnsCorrectValues()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TempBetterArray As BetterArray
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim SourceString As String
    Dim TestResult As Boolean
    
    Set TempBetterArray = New BetterArray
    Expected = Gen.GetArray(AG_BYTE, AG_JAGGED, Depth:=3)
    TempBetterArray.Items = Expected
    SourceString = TempBetterArray.ToString()
    
    'Act:
    Actual = SUT.ParseFromString(SourceString).Items
    
    ' can't use Assert.SequenceEquals due to type comparison - Bytes Will be Long in actual
    ' also, Assert.SeqenceEquals doesn't support jagged arrays: https://github.com/rubberduck-vba/Rubberduck/issues/5161
    TestResult = SequenceEquals_JaggedArray(Expected, Actual)
    
    
    'Assert:
    Assert.IsTrue TestResult, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_ParseFromString")
Private Sub ParseFromString_Jagged5DeepArrayFromToString_ReturnsCorrectValues()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TempBetterArray As BetterArray
    Dim Expected As String
    Dim Actual As String
    
    Set TempBetterArray = New BetterArray
    TempBetterArray.Items = Gen.GetArray(AG_BYTE, AG_JAGGED, 5)
    Expected = TempBetterArray.ToString()
    
    'Act:
    Actual = SUT.ParseFromString(Expected).ToString
        
    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


''''''''''''''''''''''''''''
' Method - Flatten '
''''''''''''''''''''''''''''

'@TestMethod("BetterArray_Flatten")
Private Sub Flatten_OneDimArray_ReturnsSame()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
        
    Expected = Gen.GetArray
    SUT.Items = Expected
    
    'Act:
    Actual = SUT.Flatten.Items
        
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Flatten")
Private Sub Flatten_MultiDimArray_ReturnsFlattenned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected(1 To 4) As Variant
    Dim Actual() As Variant
    Dim TestArray(1 To 2, 1 To 2) As Variant
    TestArray(1, 1) = "Foo"
    TestArray(1, 2) = "Bar"
    TestArray(2, 1) = "Fizz"
    TestArray(2, 2) = "Buzz"
    
    Expected(1) = "Foo"
    Expected(2) = "Bar"
    Expected(3) = "Fizz"
    Expected(4) = "Buzz"
    
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.Flatten.Items
        
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Flatten")
Private Sub Flatten_JaggedArray_ReturnsFlattenned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected(0 To 3) As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    TestArray = Array(Array("Foo", "Bar"), Array("Fizz", "Buzz"))
    
    Expected(0) = "Foo"
    Expected(1) = "Bar"
    Expected(2) = "Fizz"
    Expected(3) = "Buzz"
    
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.Flatten.Items
        
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Flatten")
Private Sub Flatten_EmptyInternal_ReturnsArraySizeOne()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected(0) As Variant
    Dim Actual() As Variant
    Expected(0) = Empty
    
    'Act:
    Actual = SUT.Flatten.Items
        
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

''''''''''''''''''
' Method - Clear '
''''''''''''''''''

'@TestMethod("BetterArray_Clear")
Private Sub Clear_OneDimArray_Clears()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected(0) As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim ExpectedCapacity As Long
    Dim ActualCapacity As Long
    
    Expected(0) = Empty
    TestArray = Gen.GetArray
    SUT.Items = TestArray
    ExpectedCapacity = SUT.Capacity
    'Act:
    Actual = SUT.Clear.Items
    ActualCapacity = SUT.Capacity
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
    Assert.AreEqual ExpectedCapacity, ActualCapacity, "Actual capacity <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'''''''''''''''''''''''''''
' Method - ResetToDefault '
'''''''''''''''''''''''''''

'@TestMethod("BetterArray_ResetToDefault")
Private Sub ResetToDefault_OneDimArray_Resets()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected(0) As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim ExpectedCapacity As Long
    Dim ActualCapacity As Long
    
    Expected(0) = Empty
    TestArray = Gen.GetArray
    SUT.Items = TestArray
    ExpectedCapacity = 4
    'Act:
    Actual = SUT.ResetToDefault.Items
    ActualCapacity = SUT.Capacity
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
    Assert.AreEqual ExpectedCapacity, ActualCapacity, "Actual capacity <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


''''''''''''''''''
' Method - Clone '
''''''''''''''''''

'@TestMethod("BetterArray_Clone")
Private Sub Clone_OneDimArray_CloneIsNotOriginalItemsAreSame()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim ClonedSUT As BetterArray
        
    Expected = Gen.GetArray
    SUT.Items = Expected
    
    'Act:
    Set ClonedSUT = SUT.Clone
    Actual = ClonedSUT.Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
    Assert.AreNotSame SUT, ClonedSUT, "Clone is same as original"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'''''''''''''''''''''''''''
' Method - ExtractSegment '
'''''''''''''''''''''''''''

'@TestMethod("BetterArray_ExtractSegment")
Private Sub ExtractSegment_OneDimArrayNoArgs_FullArrayReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
        
    Expected = Gen.GetArray
    SUT.Items = Expected
    
    'Act:
    Actual = SUT.ExtractSegment()
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_ExtractSegment")
Private Sub ExtractSegment_OneDimArrayJustRowArg_ArrayReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim RowIndex As Long
        
    TestArray = Gen.GetArray
    SUT.Items = TestArray
    RowIndex = 2
    Expected = Array(TestArray(RowIndex))
    
    'Act:
    Actual = SUT.ExtractSegment(RowIndex)
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_ExtractSegment")
Private Sub ExtractSegment_OneDimArrayJustColArg_ArrayReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim ColumnIndex As Long
        
    TestArray = Gen.GetArray
    SUT.Items = TestArray
    ColumnIndex = 3
    Expected = Array(TestArray(ColumnIndex))
    
    'Act:
    Actual = SUT.ExtractSegment(, ColumnIndex)
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_ExtractSegment")
Private Sub ExtractSegment_OneDimArrayRowAndColArgs_ArrayReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim RowIndex As Long
    Dim ColumnIndex As Long
        
    TestArray = Gen.GetArray
    SUT.Items = TestArray
    RowIndex = 2
    ColumnIndex = 3
    Expected = Array(TestArray(RowIndex))
    
    'Act:
    Actual = SUT.ExtractSegment(RowIndex, ColumnIndex)
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_ExtractSegment")
Private Sub ExtractSegment_JaggedArrayNoArgs_FullArrayReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
        
    Expected = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.Items = Expected
    
    'Act:
    Actual = SUT.ExtractSegment()
    
    'Assert:
    Assert.IsTrue SequenceEquals_JaggedArray(Expected, Actual), "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_ExtractSegment")
Private Sub ExtractSegment_JaggedArrayJustRowArg_ArrayReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim RowIndex As Long
        
    TestArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.Items = TestArray
    RowIndex = 2
    Expected = TestArray(RowIndex)
    
    'Act:
    Actual = SUT.ExtractSegment(RowIndex)
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_ExtractSegment")
Private Sub ExtractSegment_JaggedArrayJustColArg_ArrayReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim ColumnIndex As Long
    Dim i As Long
        
    TestArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.Items = TestArray
    ColumnIndex = 3
    ReDim Expected(LBound(TestArray) To UBound(TestArray))
    For i = LBound(Expected) To UBound(Expected)
        Expected(i) = TestArray(i)(ColumnIndex)
    Next
    
    'Act:
    Actual = SUT.ExtractSegment(, ColumnIndex)
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_ExtractSegment")
Private Sub ExtractSegment_JaggedDimArrayRowAndColArgs_ArrayReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim RowIndex As Long
    Dim ColumnIndex As Long
        
    TestArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.Items = TestArray
    RowIndex = 2
    ColumnIndex = 3
    Expected = Array(TestArray(RowIndex)(ColumnIndex))
    
    'Act:
    Actual = SUT.ExtractSegment(RowIndex, ColumnIndex)
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_ExtractSegment")
Private Sub ExtractSegment_MultiDimArrayNoArgs_FullArrayReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
        
    Expected = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.Items = Expected
    
    'Act:
    Actual = SUT.ExtractSegment()
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_ExtractSegment")
Private Sub ExtractSegment_MultiDimArrayJustRowArg_ArrayReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim RowIndex As Long
    Dim i As Long
        
    TestArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.Items = TestArray
    RowIndex = 2
    ReDim Expected(LBound(TestArray, 2) To UBound(TestArray, 2))
    For i = LBound(Expected) To UBound(Expected)
        Expected(i) = TestArray(RowIndex, i)
    Next
    
    'Act:
    Actual = SUT.ExtractSegment(RowIndex)
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_ExtractSegment")
Private Sub ExtractSegment_MultiDimArrayJustColArg_ArrayReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim ColumnIndex As Long
    Dim i As Long
        
    TestArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.Items = TestArray
    ColumnIndex = 3
    ReDim Expected(LBound(TestArray) To UBound(TestArray))
    For i = LBound(Expected) To UBound(Expected)
        Expected(i) = TestArray(i, ColumnIndex)
    Next
    
    'Act:
    Actual = SUT.ExtractSegment(, ColumnIndex)
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_ExtractSegment")
Private Sub ExtractSegment_MultiDimDimArrayRowAndColArgs_ArrayReturned()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim RowIndex As Long
    Dim ColumnIndex As Long
        
    TestArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.Items = TestArray
    RowIndex = 2
    ColumnIndex = 3
    Expected = Array(TestArray(RowIndex, ColumnIndex))
    
    'Act:
    Actual = SUT.ExtractSegment(RowIndex, ColumnIndex)
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

''''''''''''''''''''''
' Method - Transpose '
''''''''''''''''''''''
'@TestMethod("BetterArray_Transpose")
Private Sub Transpose_OneDimArray_ArrayTransposed()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim i As Long
        
    TestArray = Gen.GetArray()
    SUT.Items = TestArray
    
        
    ReDim Expected(LBound(TestArray) To UBound(TestArray), _
        LBound(TestArray) To LBound(TestArray))
    For i = LBound(TestArray) To UBound(TestArray)
        Expected(i, LBound(TestArray)) = TestArray(i)
    Next
    
    'Act:
    Actual = SUT.Transpose.Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Transpose")
Private Sub Transpose_JaggedArray_ArrayTransposed()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim Nested() As Variant
    Dim i As Long
    Dim j As Long
        
    TestArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.Items = TestArray
    
    ReDim Expected(0 To TEST_ARRAY_LENGTH - 1)

    For i = LBound(TestArray) To UBound(TestArray)
        ReDim Nested(0 To TEST_ARRAY_LENGTH - 1)
        For j = LBound(TestArray(i)) To UBound(TestArray(i))
            Nested(j) = TestArray(j)(i)
        Next
        Expected(i) = Nested
    Next
'
    'Act:
    Actual = SUT.Transpose.Items
    
    'Assert:
    Assert.IsTrue SequenceEquals_JaggedArray(Expected, Actual), "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Transpose")
Private Sub Transpose_MultiDimArray_ArrayTransposed()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim i As Long
    Dim j As Long
        
    TestArray = Gen.GetArray(ArrayType:=AG_MULTIDIMENSION)
    SUT.Items = TestArray
    
    ReDim Expected(LBound(TestArray, 2) To UBound(TestArray, 2), _
        LBound(TestArray, 1) To UBound(TestArray, 1))
    
    For i = LBound(TestArray, 1) To UBound(TestArray, 1)
        For j = LBound(TestArray, 2) To UBound(TestArray, 2)
            Expected(j, i) = TestArray(i, j)
        Next
    Next
    
    'Act:
    Actual = SUT.Transpose.Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

''''''''''''''''''''''
' Method - IndexOf '
''''''''''''''''''''''

'@TestMethod("BetterArray_IndexOf")
Private Sub IndexOf_OneDimArrayValueExists_ReturnsCorrectIndex()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Long
    Dim Actual As Long
    Dim TestArray() As Variant
    
    Expected = 3
        
    TestArray = Gen.GetArray()
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.IndexOf(TestArray(Expected))
    
    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_IndexOf")
Private Sub IndexOf_OneDimArrayValueExistsLikeComparison_ReturnsCorrectIndex()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Long
    Dim Actual As Long
    Dim TestArray() As Variant
    Dim Pattern As String
    
    Expected = 3
    Pattern = "a*a"
    TestArray = Array("Zero", "One", "Two", "aBBBa")
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.IndexOf(Pattern, , CT_LIKENESS)
    
    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_IndexOf")
Private Sub IndexOf_OneDimArrayLikeComparisonPatternNotString_ThrowsError()
    Const ExpectedError As Long = ErrorCodes.EC_STRING_TYPE_EXPECTED
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore VariableNotUsed
    Dim Expected As Long
    '@Ignore VariableNotUsed
    Dim Actual As Long
    Dim TestArray() As Variant
    Dim Pattern As Collection
    
    Expected = 3
    Set Pattern = New Collection
    TestArray = Array("Zero", "One", "Two", "aBBBa")
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.IndexOf(Pattern, , CT_LIKENESS)
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("BetterArray_IndexOf")
Private Sub IndexOf_OneDimArrayValueMissing_ReturnsMISSING_LONG()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Long
    Dim Actual As Long
    Dim TestArray() As Variant
    
    Expected = MISSING_LONG
        
    TestArray = Gen.GetArray()
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.IndexOf("Foo")
    
    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_IndexOf")
Private Sub IndexOf_JaggedArray_ReturnsCorrectIndex()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Long
    Dim Actual As Long
    Dim TestArray() As Variant
    
    Expected = 3
        
    TestArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.IndexOf(TestArray(Expected))
    
    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

''''''''''''''''''''''
' Method - Unique '
''''''''''''''''''''''

'@TestMethod("BetterArray_Unique")
Private Sub Unique_OneDimArray_ReturnsUniqueList()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    
    TestArray = Array(1, 2, 2, 1, 3, 4, 5, 5, 6, 3)
    Expected = Array(1, 2, 3, 4, 5, 6)
        
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.Unique.Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Unique")
Private Sub Unique_JaggedArray_ReturnsUniqueList()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    
    TestArray = Array( _
        Array(1, 2, 3), _
        Array("Foo", "Bar"), _
        Array("Foo", "Fizz"), _
        Array(1, 2, 3), _
        Array("Foo", "Bar") _
    )
    Expected = Array( _
        Array(1, 2, 3), _
        Array("Foo", "Bar"), _
        Array("Foo", "Fizz") _
    )
        
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.Unique.Items
    
    'Assert:
    Assert.IsTrue SequenceEquals_JaggedArray(Expected, Actual), "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Unique")
Private Sub Unique_JaggedArrayColumnIndexBase0_ReturnsUniqueList()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    
    TestArray = Array( _
        Array(1, "Foo", 3), _
        Array(1, "Bar", 3), _
        Array(1, "Foo", 3), _
        Array(1, "Foo", 3), _
        Array(1, "Bar", 3), _
        Array(1, "Fizz", 3), _
        Array(1, "Fizz", 3), _
        Array(1, "Buzz", 3) _
    )
    Expected = Array( _
        Array(1, "Foo", 3), _
        Array(1, "Bar", 3), _
        Array(1, "Fizz", 3), _
        Array(1, "Buzz", 3) _
    )
        
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.Unique(2).Items
    
    'Assert:
    Assert.IsTrue SequenceEquals_JaggedArray(Expected, Actual), "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Unique")
Private Sub Unique_JaggedArrayColumnIndexBaseNegativeBase_ReturnsUniqueList()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    
    TestArray = Array( _
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
    ReDim Expected(-10 To -7)
    
    
    Expected(-10) = Array(1, "Foo", 3)
    Expected(-9) = Array(1, "Bar", 3)
    Expected(-8) = Array(1, "Fizz", 3)
    Expected(-7) = Array(1, "Buzz", 3)
    
    SUT.LowerBound = -10
    
    ' rebase nested items
    SUT.Items = Expected
    Expected = SUT.Items

    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.Unique(2).Items
    
    'Assert:
    Assert.IsTrue SequenceEquals_JaggedArray(Expected, Actual), "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Unique")
Private Sub Unique_JaggedArrayColumnIndexPositiveBase_ReturnsUniqueList()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    
    TestArray = Array( _
        Array(1, "Foo", 3), _
        Array(1, "Bar", 3), _
        Array(1, "Foo", 3), _
        Array(1, "Foo", 3), _
        Array(1, "Bar", 3), _
        Array(1, "Fizz", 3), _
        Array(1, "Fizz", 3), _
        Array(1, "Buzz", 3) _
    )
    
    
    ReDim Expected(10 To 13)
    Expected(10) = Array(1, "Foo", 3)
    Expected(11) = Array(1, "Bar", 3)
    Expected(12) = Array(1, "Fizz", 3)
    Expected(13) = Array(1, "Buzz", 3)
    
    SUT.LowerBound = 10
    
    ' rebase expected itms
    SUT.Items = Expected
    Expected = SUT.Items
    
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.Unique(2).Items
    
    'Assert:
    Assert.IsTrue SequenceEquals_JaggedArray(Expected, Actual), "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'''''''''''''''''''
' Method - Remove '
'''''''''''''''''''

'@TestMethod("BetterArray_Remove")
Private Sub Remove_OneDimArray_RemovesElementAtIndex()
    On Error GoTo TestFail
    
    'Arrange:
    Const RemoveIndex As Long = 2
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim ExpectedLength As Long
    Dim ActualLength As Long
    
    TestArray = Array("Foo", "Bar", "Fizz", "Buzz")
    Expected = Array("Foo", "Bar", "Buzz")
    ExpectedLength = Gen.GetArrayLength(Expected)
    
    SUT.Items = TestArray
    
    'Act:
    
    ActualLength = SUT.Remove(RemoveIndex)
    Actual = SUT.Items
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
    Assert.AreEqual ExpectedLength, ActualLength, "Actual length <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Remove")
Private Sub Remove_JaggedArray_RemovesElementAtIndex()
    On Error GoTo TestFail
    
    'Arrange:
    Const RemoveIndex As Long = 2
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim ExpectedLength As Long
    Dim ActualLength As Long
    
    TestArray = Array( _
        Array(1, 2, 3), _
        Array("Foo", "Bar"), _
        Array("Foo", "Fizz"), _
        Array(1, 2, 3), _
        Array("Foo", "Bar") _
    )
    Expected = Array( _
        Array(1, 2, 3), _
        Array("Foo", "Bar"), _
        Array(1, 2, 3), _
        Array("Foo", "Bar") _
    )
    ExpectedLength = Gen.GetArrayLength(Expected)
    
    SUT.Items = TestArray
    
    'Act:
    
    ActualLength = SUT.Remove(RemoveIndex)
    Actual = SUT.Items
    'Assert:
    Assert.IsTrue SequenceEquals_JaggedArray(Expected, Actual), "Actual <> expected"
    Assert.AreEqual ExpectedLength, ActualLength, "Actual length <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Remove")
Private Sub Remove_MultiDimArray_RemovesElementAtIndex()
    On Error GoTo TestFail
    
    'Arrange:
    Const RemoveIndex As Long = 2
    Dim Expected(1 To 2, 1 To 2) As Variant
    Dim Actual() As Variant
    Dim TestArray(1 To 3, 1 To 2) As Variant
    Dim ExpectedLength As Long
    Dim ActualLength As Long
    
    TestArray(1, 1) = "Foo"
    TestArray(1, 2) = "Bar"
    TestArray(2, 1) = "Fizz"
    TestArray(2, 2) = "Buzz"
    TestArray(3, 1) = "Whizz"
    TestArray(3, 2) = "Bang"
    
    Expected(1, 1) = "Foo"
    Expected(1, 2) = "Bar"
    Expected(2, 1) = "Whizz"
    Expected(2, 2) = "Bang"

    ExpectedLength = Gen.GetArrayLength(Expected)
    
    SUT.Items = TestArray
    
    'Act:
    
    ActualLength = SUT.Remove(RemoveIndex)
    Actual = SUT.Items
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
    Assert.AreEqual ExpectedLength, ActualLength, "Actual length <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Remove")
Private Sub Remove_OneDimArrayRemoveFirst_RemovesElementAtIndex()
    On Error GoTo TestFail
    
    'Arrange:
    Const RemoveIndex As Long = 0
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim ExpectedLength As Long
    Dim ActualLength As Long
    
    TestArray = Array("Foo", "Bar", "Fizz", "Buzz")
    Expected = Array("Bar", "Fizz", "Buzz")
    ExpectedLength = Gen.GetArrayLength(Expected)
    
    SUT.Items = TestArray
    
    'Act:
    
    ActualLength = SUT.Remove(RemoveIndex)
    Actual = SUT.Items
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
    Assert.AreEqual ExpectedLength, ActualLength, "Actual length <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Remove")
Private Sub Remove_OneDimArrayRemoveLast_RemovesElementAtIndex()
    On Error GoTo TestFail
    
    'Arrange:
    Const RemoveIndex As Long = 3
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim ExpectedLength As Long
    Dim ActualLength As Long
    
    TestArray = Array("Foo", "Bar", "Fizz", "Buzz")
    Expected = Array("Foo", "Bar", "Fizz")
    ExpectedLength = Gen.GetArrayLength(Expected)
    
    SUT.Items = TestArray
    
    'Act:
    
    ActualLength = SUT.Remove(RemoveIndex)
    Actual = SUT.Items
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
    Assert.AreEqual ExpectedLength, ActualLength, "Actual length <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Remove")
Private Sub Remove_OneDimArrayIndexExceedsBounds_NothingRemoved()
    On Error GoTo TestFail
    
    'Arrange:
    Const RemoveIndex As Long = 100
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim ExpectedLength As Long
    Dim ActualLength As Long
    
    TestArray = Array("Foo", "Bar", "Fizz", "Buzz")
    Expected = Array("Foo", "Bar", "Fizz", "Buzz")
    ExpectedLength = Gen.GetArrayLength(Expected)
    
    SUT.Items = TestArray
    
    'Act:
    
    ActualLength = SUT.Remove(RemoveIndex)
    Actual = SUT.Items
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
    Assert.AreEqual ExpectedLength, ActualLength, "Actual length <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

''''''''''''''''''
' Method - Every '
''''''''''''''''''

'@TestMethod("BetterArray_Every")
Private Sub Every_OneDimArrayOfSameString_ReturnsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    TestArray = Array("Foo", "Foo", "Foo", "Foo")
    
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.Every("Foo")
    
    'Assert:
    Assert.IsTrue Actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Every")
Private Sub Every_OneDimArrayOfDifferentString_ReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    TestArray = Array("Foo", "Bar", "Foo", "Foo")
    
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.Every("Foo")
    
    'Assert:
    Assert.IsFalse Actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Every")
Private Sub Every_JaggedDimArrayOfSameString_ReturnsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    TestArray = Array(Array("Foo", "Foo", "Foo", "Foo"), Array("Foo", "Foo", "Foo", "Foo"))
    
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.Every("Foo")
    
    'Assert:
    Assert.IsTrue Actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Every")
Private Sub Every_JaggedDimArrayOfSameString_ReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    TestArray = Array(Array("Foo", "Bar", "Foo", "Foo"), Array("Foo", "Foo", "Foo", "Foo"))
    
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.Every("Foo")
    
    'Assert:
    Assert.IsFalse Actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Every")
Private Sub Every_MiltiDimArrayOfSameString_ReturnsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    ReDim TestArray(1 To 2, 1 To 2)
    TestArray(1, 1) = "Foo"
    TestArray(1, 2) = "Foo"
    TestArray(2, 1) = "Foo"
    TestArray(2, 2) = "Foo"
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.Every("Foo")
    
    'Assert:
    Assert.IsTrue Actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Every")
Private Sub Every_MiltiDimArrayOfDifferentString_ReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    ReDim TestArray(1 To 2, 1 To 2)
    TestArray(1, 1) = "Foo"
    TestArray(1, 2) = "Bar"
    TestArray(2, 1) = "Foo"
    TestArray(2, 2) = "Foo"
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.Every("Foo")
    
    'Assert:
    Assert.IsFalse Actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'''''''''''''''''''''
' Method - EveryType'
'''''''''''''''''''''

'@TestMethod("BetterArray_EveryType")
Private Sub EveryType_OneDimArrayOfSameString_ReturnsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    TestArray = Array("Foo", "Foo", "Foo", "Foo")
    
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.EveryType("string")
    
    'Assert:
    Assert.IsTrue Actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_EveryType")
Private Sub EveryType_OneDimArrayOfDifferentTypes_ReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    TestArray = Array("Foo", 1, 1.2, "Foo")
    
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.EveryType("string")
    
    'Assert:
    Assert.IsFalse Actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_EveryType")
Private Sub EveryType_JaggedDimArrayOfSameString_ReturnsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    TestArray = Array( _
        Array("Foo", "Foo", "Foo", "Foo"), _
        Array("Foo", "Foo", "Foo", "Foo") _
    )
    
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.EveryType("string")
    
    'Assert:
    Assert.IsTrue Actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_EveryType")
Private Sub EveryType_JaggedDimArrayOfSameString_ReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    TestArray = Array( _
        Array("Foo", 1.123, "Foo", "Foo"), _
        Array("Foo", "Foo", "Foo", "Foo") _
    )
    
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.EveryType("string")
    
    'Assert:
    Assert.IsFalse Actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_EveryType")
Private Sub EveryType_MiltiDimArrayOfSameString_ReturnsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    ReDim TestArray(1 To 2, 1 To 2)
    TestArray(1, 1) = "Foo"
    TestArray(1, 2) = "Foo"
    TestArray(2, 1) = "Foo"
    TestArray(2, 2) = "Foo"
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.EveryType("string")
    
    'Assert:
    Assert.IsTrue Actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_EveryType")
Private Sub EveryType_MiltiDimArrayOfDifferentString_ReturnsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual As Boolean
    Dim TestArray() As Variant
    
    ReDim TestArray(1 To 2, 1 To 2)
    TestArray(1, 1) = "Foo"
    TestArray(1, 2) = 1.123
    TestArray(2, 1) = "Foo"
    TestArray(2, 2) = "Foo"
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.EveryType("string")
    
    'Assert:
    Assert.IsFalse Actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


''''''''''''''''''''''
' Method - Fill  '
''''''''''''''''''''''

'@TestMethod("BetterArray_Fill")
Private Sub Fill_OneDimArray2To4_SpecifiedIndicesFilled()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual() As Variant
    Dim Expected() As Variant
    
    Dim TestArray() As Variant
    
    TestArray = Gen.GetArray
    
    Const FillVal As Long = 0
        
    Expected = TestArray
    Dim i As Long
    For i = 2 To 4
        Expected(i) = FillVal
    Next
    
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.Fill(FillVal, 2, 4).Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Fill")
Private Sub Fill_OneDimArray1ToEnd_SpecifiedIndicesFilled()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual() As Variant
    Dim Expected() As Variant
    
    Dim TestArray() As Variant
    
    TestArray = Gen.GetArray
    
    Const FillVal As Long = 5
        
    Expected = TestArray
    Dim i As Long
    For i = 1 To UBound(Expected)
        Expected(i) = FillVal
    Next
    
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.Fill(FillVal, 1).Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Fill")
Private Sub Fill_OneDimArrayAll_SpecifiedIndicesFilled()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Actual() As Variant
    Dim Expected() As Variant
    
    Dim TestArray() As Variant
    
    TestArray = Gen.GetArray
    
    Const FillVal As Long = 6
        
    Expected = TestArray
    Dim i As Long
    For i = LBound(Expected) To UBound(Expected)
        Expected(i) = FillVal
    Next
    
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.Fill(FillVal).Items
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

''''''''''''''''''''''
' Method - LastIndexOf '
''''''''''''''''''''''

'@TestMethod("BetterArray_LastIndexOf")
Private Sub LastIndexOf_OneDimArrayValueExists_ReturnsCorrectIndex()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Long
    Dim Actual As Long
    Dim TestArray() As Variant
    
    Expected = 3
        
    TestArray = Array("Dodo", "Tiger", "Penguin", "Dodo")
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.LastIndexOf("Dodo")
    
    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_LastIndexOf")
Private Sub LastIndexOf_OneDimArrayValueExistsLikeComparison_ReturnsCorrectIndex()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Long
    Dim Actual As Long
    Dim TestArray() As Variant
    Dim Pattern As String
    
    Expected = 3
    Pattern = "a*a"
    TestArray = Array("Zero", "One", "Two", "aBBBa")
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.IndexOf(Pattern, , CT_LIKENESS)
    
    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_LastIndexOf")
Private Sub LastIndexOf_OneDimArrayLikeComparisonPatternNotString_ThrowsError()
    Const ExpectedError As Long = ErrorCodes.EC_STRING_TYPE_EXPECTED
    On Error GoTo TestFail
    
    'Arrange:
    '@Ignore VariableNotUsed
    Dim Expected As Long
    '@Ignore VariableNotUsed
    Dim Actual As Long
    Dim TestArray() As Variant
    Dim Pattern As Collection
    
    Expected = 3
    Set Pattern = New Collection
    TestArray = Array("Zero", "One", "Two", "aBBBa")
    SUT.Items = TestArray
    
    'Act:
    Actual = SUT.IndexOf(Pattern, , CT_LIKENESS)
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("BetterArray_LastIndexOf")
Private Sub LastIndexOf_OneDimArrayValueMissing_ReturnsMISSING_LONG()
    On Error GoTo TestFail

    'Arrange:
    Dim Expected As Long
    Dim Actual As Long
    Dim TestArray() As Variant

    Expected = MISSING_LONG

    TestArray = Gen.GetArray()
    SUT.Items = TestArray

    'Act:
    Actual = SUT.LastIndexOf("Foo")

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_LastIndexOf")
Private Sub LastIndexOf_JaggedArray_ReturnsCorrectIndex()
    On Error GoTo TestFail

    'Arrange:
    Dim Expected As Long
    Dim Actual As Long
    Dim TestArray() As Variant

    Expected = 3

    TestArray = Gen.GetArray(ArrayType:=AG_JAGGED)
    SUT.Items = TestArray

    'Act:
    Actual = SUT.LastIndexOf(TestArray(Expected))

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'''''''''''''''''''
' Method - Splice '
'''''''''''''''''''

'@TestMethod("BetterArray_Splice")
Private Sub Splice_OneDimArrayInsertAtIndex1_Success()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim ActualResult() As Variant
    Dim ExpectedResult() As Variant
    
    Expected = Array("Jan", "Feb", "March", "April", "June")
    TestArray = Array("Jan", "March", "April", "June")
    SUT.Items = TestArray
    ReDim ExpectedResult(0)

    'Act:
    ActualResult = SUT.Splice(1, 0, "Feb")
    Actual = SUT.Items

    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
    Assert.SequenceEquals ExpectedResult, ActualResult, "ActualResult <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Splice")
Private Sub Splice_OneDimArrayInsertAtIndex1Delete1_Success()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim ActualResult() As Variant
    Dim ExpectedResult() As Variant
    

    Expected = Array("Jan", "Feb", "March", "April", "May")
    TestArray = Array("Jan", "Feb", "March", "April", "June")
    SUT.Items = TestArray
    ExpectedResult = Array("June")
    
    'Act:
    ActualResult = SUT.Splice(4, 1, "May")
    Actual = SUT.Items

    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
    Assert.SequenceEquals ExpectedResult, ActualResult, "ActualResult <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Splice")
Private Sub Splice_OneDimArrayInsertAtIndex2Delete0Insert2_Success()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim ActualResult() As Variant
    Dim ExpectedResult() As Variant
    

    Expected = Array("Banana", "Orange", "Lemon", "Kiwi", "Apple", "Mango")
    TestArray = Array("Banana", "Orange", "Apple", "Mango")
    SUT.Items = TestArray
    ReDim ExpectedResult(0)
    
    'Act:
    ActualResult = SUT.Splice(2, 0, "Lemon", "Kiwi")
    Actual = SUT.Items

    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
    Assert.SequenceEquals ExpectedResult, ActualResult, "ActualResult <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Splice")
Private Sub Splice_OneDimArrayInsertAtIndex2Delete1Insert2_Success()
    On Error GoTo TestFail
        
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim ActualResult() As Variant
    Dim ExpectedResult() As Variant
    

    Expected = Array("Banana", "Orange", "Lemon", "Kiwi", "Mango")
    TestArray = Array("Banana", "Orange", "Apple", "Mango")
    SUT.Items = TestArray
    ExpectedResult = Array("Apple")
    
    'Act:
    ActualResult = SUT.Splice(2, 1, "Lemon", "Kiwi")
    Actual = SUT.Items

    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
    Assert.SequenceEquals ExpectedResult, ActualResult, "ActualResult <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Splice")
Private Sub Splice_OneDimArrayInsertAtIndex2Delete2Insert0_Success()
    On Error GoTo TestFail
        
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim ActualResult() As Variant
    Dim ExpectedResult() As Variant
    

    Expected = Array("Banana", "Orange", "Kiwi")
    TestArray = Array("Banana", "Orange", "Apple", "Mango", "Kiwi")
    SUT.Items = TestArray
    ExpectedResult = Array("Apple", "Mango")
    
    'Act:
    ActualResult = SUT.Splice(2, 2)
    Actual = SUT.Items

    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
    Assert.SequenceEquals ExpectedResult, ActualResult, "ActualResult <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_Splice")
Private Sub Splice_OneDimArrayBase1InsertAtIndex2Delete0Insert2_Success()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim ActualResult() As Variant
    Dim ExpectedResult() As Variant
        
    SUT.LowerBound = 1
    
    ReDim Expected(1 To 6)
    Expected(1) = "Banana"
    Expected(2) = "Orange"
    Expected(3) = "Lemon"
    Expected(4) = "Kiwi"
    Expected(5) = "Apple"
    Expected(6) = "Mango"
    
    TestArray = Array("Banana", "Orange", "Apple", "Mango")
    SUT.Items = TestArray
    ReDim ExpectedResult(0)
    
    'Act:
    ActualResult = SUT.Splice(3, 0, "Lemon", "Kiwi")
    Actual = SUT.Items

    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
    Assert.SequenceEquals ExpectedResult, ActualResult, "ActualResult <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Splice")
Private Sub Splice_OneDimArrayBase1InsertAtIndex2Delete1Insert2_Success()
    On Error GoTo TestFail
        
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim ActualResult() As Variant
    Dim ExpectedResult() As Variant
    
    SUT.LowerBound = 1
    ReDim Expected(1 To 5)
    Expected(1) = "Banana"
    Expected(2) = "Orange"
    Expected(3) = "Lemon"
    Expected(4) = "Kiwi"
    Expected(5) = "Mango"
    
    TestArray = Array("Banana", "Orange", "Apple", "Mango")
    SUT.Items = TestArray
    ExpectedResult = Array("Apple")
    
    'Act:
    ActualResult = SUT.Splice(3, 1, "Lemon", "Kiwi")
    Actual = SUT.Items

    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
    Assert.SequenceEquals ExpectedResult, ActualResult, "ActualResult <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Splice")
Private Sub Splice_OneDimArrayBase1InsertAtIndex2Delete2Insert0_Success()
    On Error GoTo TestFail
        
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim ActualResult() As Variant
    Dim ExpectedResult() As Variant
    
    SUT.LowerBound = 1
    
    ReDim Expected(1 To 3)
    Expected(1) = "Banana"
    Expected(2) = "Orange"
    Expected(3) = "Kiwi"
    
    TestArray = Array("Banana", "Orange", "Apple", "Mango", "Kiwi")
    SUT.Items = TestArray
    ExpectedResult = Array("Apple", "Mango")
    
    'Act:
    ActualResult = SUT.Splice(3, 2)
    Actual = SUT.Items

    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
    Assert.SequenceEquals ExpectedResult, ActualResult, "ActualResult <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Splice")
Private Sub Splice_OneDimArrayBaseNegative1InsertAtIndex2Delete0Insert2_Success()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim ActualResult() As Variant
    Dim ExpectedResult() As Variant
    
    SUT.LowerBound = -1
    
    ReDim Expected(-1 To 4)
    Expected(-1) = "Banana"
    Expected(0) = "Orange"
    Expected(1) = "Lemon"
    Expected(2) = "Kiwi"
    Expected(3) = "Apple"
    Expected(4) = "Mango"

    TestArray = Array("Banana", "Orange", "Apple", "Mango")
    SUT.Items = TestArray
    ReDim ExpectedResult(0)
    
    'Act:
    ActualResult = SUT.Splice(1, 0, "Lemon", "Kiwi")
    Actual = SUT.Items

    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
    Assert.SequenceEquals ExpectedResult, ActualResult, "ActualResult <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Splice")
Private Sub Splice_OneDimArrayBaseNegative1InsertAtIndex2Delete1Insert2_Success()
    On Error GoTo TestFail
        
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim ActualResult() As Variant
    Dim ExpectedResult() As Variant
    
    SUT.LowerBound = -1
    ReDim Expected(-1 To 3)
    Expected(-1) = "Banana"
    Expected(0) = "Orange"
    Expected(1) = "Lemon"
    Expected(2) = "Kiwi"
    Expected(3) = "Mango"
    
    TestArray = Array("Banana", "Orange", "Apple", "Mango")
    SUT.Items = TestArray
    ExpectedResult = Array("Apple")
    
    'Act:
    ActualResult = SUT.Splice(1, 1, "Lemon", "Kiwi")
    Actual = SUT.Items

    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
    Assert.SequenceEquals ExpectedResult, ActualResult, "ActualResult <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_Splice")
Private Sub Splice_OneDimArrayBaseNegative1InsertAtIndex2Delete2Insert0_Success()
    On Error GoTo TestFail
        
    'Arrange:
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim TestArray() As Variant
    Dim ActualResult() As Variant
    Dim ExpectedResult() As Variant
    
    SUT.LowerBound = -1
    ReDim Expected(-1 To 1)
    Expected(-1) = "Banana"
    Expected(0) = "Orange"
    Expected(1) = "Kiwi"
    
    TestArray = Array("Banana", "Orange", "Apple", "Mango", "Kiwi")
    SUT.Items = TestArray
    ExpectedResult = Array("Apple", "Mango")
    
    'Act:
    ActualResult = SUT.Splice(1, 2)
    Actual = SUT.Items

    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> expected"
    Assert.SequenceEquals ExpectedResult, ActualResult, "ActualResult <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
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

    Dim Expected() As Variant
    Dim Actual() As Variant
    ReDim Expected(0 To 9)

    Expected(0) = Array("Region", "Country A", "Item Type", "Sales Channel", "Order Priority", "Order Date", "Order ID", "Ship Date", "Units Sold", "Unit Price", "Unit Cost", "Total Revenue", "Total Cost", "Total Profit")
    Expected(1) = Array("Sub-Saharan Africa", "Chad", "Office Supplies", "Online", "L", "1/27/2011", "292494523", "2/12/2011", "4484", "651.21", "524.96", "2920025.64", "2353920.64", "566105.00")
    Expected(2) = Array("Europe", "Latvia", "Beverages", "Online", "C", "12 / 28 / 2015", "361825549", "1 / 23 / 2016", "1075", "47.45", "31.79", "51008.75", "34174.25", "16834.5")
    Expected(3) = Array("Middle East and North Africa", "Pakistan", "Vegetables", "Offline", "C", "1/13/2011", "141515767", "2/1/2011", "6515", "154.06", "90.93", "1003700.90", "592408.95", "411291.95")
    Expected(4) = Array("Sub-Saharan Africa", "Democratic Republic of the Congo", "Household", "Online", "C", "9/11/2012", "500364005", "10/6/2012", "7683", "668.27", "502.54", "5134318.41", "3861014.82", "1273303.59")
    Expected(5) = Array("Europe", "Czech Republic", "Beverages", "Online", "C", "10/27/2015", "127481591", "12/5/2015", "3491", "47.45", "31.79", "165647.95", "110978.89", "54669.06")
    Expected(6) = Array("Sub-Saharan Africa", "South Africa", "Beverages", "Offline", "H", "7/10/2012", "482292354", "8/21/2012", "9880", "47.45", "31.79", "468806.00", "314085.20", "154720.80")
    Expected(7) = Array("Asia", "Laos", "Vegetables", "Online", "L", "2 / 20 / 2011", "844532620", "3 / 20 / 2011", "4825", "154.06", "90.93", "743339.50", "438737.25", "304602.25")
    Expected(8) = Array("Asia", "China", "Baby Food", "Online", "C", "4/10/2017", "564251220", "5/12/2017", "3330", "255.28", "159.42", "850082.40", "530868.60", "319213.80")
    Expected(9) = Array("Sub-Saharan Africa", "Eritrea", "Meat", "Online", "L", "11/21/2014", "411809480", "1/10/2015", "2431", "421.89", "364.69", "1025614.59", "886561.39", "139053.20")
           
    'Act:
    Actual = SUT.FromCSVString(TEST_DATA).Items

    'Assert:
    Assert.IsTrue SequenceEquals_JaggedArray(Expected, Actual), "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_FromCSVString")
Private Sub FromCSVString_RFC4180_ReturnsJagged()
    On Error GoTo TestFail
        
    'Arrange:
    '@Ignore UseMeaningfulName
    Dim Line1 As String
    '@Ignore UseMeaningfulName
    Dim Line2 As String
    '@Ignore UseMeaningfulName
    Dim Line3 As String
    Dim CSVData As String
    Line1 = _
        WrapQuoteUtil("Field with " & vbCrLf & "multiple lines") & " ," & _
        WrapQuoteUtil("Another field " & vbCrLf & "with some " & vbCrLf & "line breaks inside") & " , " & _
        WrapQuoteUtil("Include some  comma, for test, and some [" & WrapQuoteUtil() & "] Quotes") & " , " & _
        WrapQuoteUtil("Normal field here") & vbCrLf
    Line2 = "1, 2, 3 ,4 " & vbCrLf
    Line3 = "Field 1, Field 2 , Field 3 , Field 4"
    CSVData = Line1 & Line2 & Line3

    Dim Expected() As Variant
    Dim Actual() As Variant
    ReDim Expected(0 To 2)

    Expected(0) = Array( _
        "Field with " & vbCrLf & "multiple lines", _
        "Another field " & vbCrLf & "with some " & vbCrLf & "line breaks inside", _
        "Include some  comma, for test, and some [" & WrapQuoteUtil() & "] Quotes", _
        "Normal field here" _
    )
    Expected(1) = Array("1", "2", "3", "4")
    Expected(2) = Array("Field 1", "Field 2", "Field 3", "Field 4")
           
    'Act:
    Actual = SUT.FromCSVString(CSVData).Items

    'Assert:
    Assert.IsTrue SequenceEquals_JaggedArray(Expected, Actual), "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BetterArray_FromCSVString")
Private Sub FromCSVString_NullString_ReturnsJagged()
    On Error GoTo TestFail
        
    'Arrange:
    Dim CSVData As String
    CSVData = _
        WrapQuoteUtil() & "," & vbCrLf & _
        "," & WrapQuoteUtil() & " " & vbCrLf & _
        "Field1,Field2" & vbCrLf

    Dim Expected() As Variant
    Dim Actual() As Variant
    ReDim Expected(0 To 2)

    Expected(0) = Array(vbNullString, vbNullString)
    Expected(1) = Array(vbNullString, vbNullString)
    Expected(2) = Array("Field1", "Field2")
           
    'Act:
    Actual = SUT.FromCSVString(CSVData).Items

    'Assert:
    Assert.IsTrue SequenceEquals_JaggedArray(Expected, Actual), "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'''''''''''''''
' ToCSVString '
'''''''''''''''

'@TestMethod("BetterArray_ToCSVString")
Private Sub ToCSVString_Simple10RowWithHeaders_ValidStringReturned()
    On Error GoTo TestFail

    'Arrange:
    Const Expected As String = _
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

    Dim Headers() As Variant
    Dim TestDatum() As Variant
    Dim Actual As String
    ReDim TestDatum(0 To 8)

    Headers = Array("Region", "Country", "Item Type", "Sales Channel", "Order Priority", "Order Date", "Order ID", "Ship Date", "Units Sold", "Unit Price", "Unit Cost", "Total Revenue", "Total Cost", "Total Profit")
    TestDatum(0) = Array("Sub-Saharan Africa", "Chad", "Office Supplies", "Online", "L", "1/27/2011", "292494523", "2/12/2011", "4484", "651.21", "524.96", "2920025.64", "2353920.64", "566105.00")
    TestDatum(1) = Array("Europe", "Latvia", "Beverages", "Online", "C", "12/28/2015", "361825549", "1/23/2016", "1075", "47.45", "31.79", "51008.75", "34174.25", "16834.50")
    TestDatum(2) = Array("Middle East and North Africa", "Pakistan", "Vegetables", "Offline", "C", "1/13/2011", "141515767", "2/1/2011", "6515", "154.06", "90.93", "1003700.90", "592408.95", "411291.95")
    TestDatum(3) = Array("Sub-Saharan Africa", "Democratic Republic of the Congo", "Household", "Online", "C", "9/11/2012", "500364005", "10/6/2012", "7683", "668.27", "502.54", "5134318.41", "3861014.82", "1273303.59")
    TestDatum(4) = Array("Europe", "Czech Republic", "Beverages", "Online", "C", "10/27/2015", "127481591", "12/5/2015", "3491", "47.45", "31.79", "165647.95", "110978.89", "54669.06")
    TestDatum(5) = Array("Sub-Saharan Africa", "South Africa", "Beverages", "Offline", "H", "7/10/2012", "482292354", "8/21/2012", "9880", "47.45", "31.79", "468806.00", "314085.20", "154720.80")
    TestDatum(6) = Array("Asia", "Laos", "Vegetables", "Online", "L", "2/20/2011", "844532620", "3/20/2011", "4825", "154.06", "90.93", "743339.50", "438737.25", "304602.25")
    TestDatum(7) = Array("Asia", "China", "Baby Food", "Online", "C", "4/10/2017", "564251220", "5/12/2017", "3330", "255.28", "159.42", "850082.40", "530868.60", "319213.80")
    TestDatum(8) = Array("Sub-Saharan Africa", "Eritrea", "Meat", "Online", "L", "11/21/2014", "411809480", "1/10/2015", "2431", "421.89", "364.69", "1025614.59", "886561.39", "139053.20")

    'Act:
    SUT.Items = TestDatum
    Actual = SUT.ToCSVString(Headers:=Headers)
    
    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_ToCSVString")
Private Sub ToCSVString_Simple10RowNoHeaders_ValidStringReturned()
    On Error GoTo TestFail

    'Arrange:
    Const Expected As String = _
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
        
        
    Dim TestDatum() As Variant
    Dim Actual As String
    ReDim TestDatum(0 To 9)

    TestDatum(0) = Array("Region", "Country", "Item Type", "Sales Channel", "Order Priority", "Order Date", "Order ID", "Ship Date", "Units Sold", "Unit Price", "Unit Cost", "Total Revenue", "Total Cost", "Total Profit")
    TestDatum(1) = Array("Sub-Saharan Africa", "Chad", "Office Supplies", "Online", "L", "1/27/2011", "292494523", "2/12/2011", "4484", "651.21", "524.96", "2920025.64", "2353920.64", "566105.00")
    TestDatum(2) = Array("Europe", "Latvia", "Beverages", "Online", "C", "12/28/2015", "361825549", "1/23/2016", "1075", "47.45", "31.79", 51008.75, "34174.25", "16834.50")
    TestDatum(3) = Array("Middle East and North Africa", "Pakistan", "Vegetables", "Offline", "C", "1/13/2011", "141515767", "2/1/2011", "6515", "154.06", "90.93", "1003700.90", "592408.95", "411291.95")
    TestDatum(4) = Array("Sub-Saharan Africa", "Democratic Republic of the Congo", "Household", "Online", "C", "9/11/2012", "500364005", "10/6/2012", "7683", "668.27", "502.54", "5134318.41", "3861014.82", "1273303.59")
    TestDatum(5) = Array("Europe", "Czech Republic", "Beverages", "Online", "C", "10/27/2015", "127481591", "12/5/2015", "3491", "47.45", "31.79", "165647.95", "110978.89", "54669.06")
    TestDatum(6) = Array("Sub-Saharan Africa", "South Africa", "Beverages", "Offline", "H", "7/10/2012", "482292354", "8/21/2012", "9880", "47.45", "31.79", "468806.00", "314085.20", "154720.80")
    TestDatum(7) = Array("Asia", "Laos", "Vegetables", "Online", "L", "2/20/2011", "844532620", "3/20/2011", "4825", "154.06", "90.93", "743339.50", "438737.25", "304602.25")
    TestDatum(8) = Array("Asia", "China", "Baby Food", "Online", "C", "4/10/2017", "564251220", "5/12/2017", "3330", "255.28", "159.42", "850082.40", "530868.60", "319213.80")
    TestDatum(9) = Array("Sub-Saharan Africa", "Eritrea", "Meat", "Online", "L", "11/21/2014", "411809480", "1/10/2015", "2431", "421.89", "364.69", "1025614.59", "886561.39", "139053.20")

    'Act:
    SUT.Items = TestDatum
    Actual = SUT.ToCSVString()
    
    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_ToCSVString")
Private Sub ToCSVString_1DArray_ValidStringReturned()
    On Error GoTo TestFail

    'Arrange:
    Const Expected As String = _
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

    Dim TestDatum() As Variant
    Dim Actual As String

    TestDatum = Array("Region", "Country A", "Item Type", "Sales Channel", "Order Priority", "Order Date", "Order ID", "Ship Date", "Units Sold", "Unit Price", "Unit Cost", "Total Revenue", "Total Cost", "Total Profit")

    'Act:
    SUT.Items = TestDatum
    Actual = SUT.ToCSVString()

    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_ToCSVString")
Private Sub ToCSVString_Simple2RowNoHeadersDatesAndDoubles_ValidStringReturned()
    On Error GoTo TestFail

    'Arrange:
    Const Expected As String = _
        "Sub-Saharan Africa,Chad,Office Supplies,Online,L,1/27/2011,""292,494,523.00"",12/02/2011,""4,484.00"",651.21,524.96,""2,920,025.64"",""2,353,920.64"",""566,105.00""" & vbCrLf & _
        "Europe,Latvia,Beverages,Online,C,12/28/2015,""361,825,549.00"",1/23/2016,""1,075.00"",47.45,31.79,""51,008.75"",""34,174.25"",""16,834.50"""

        
    Dim TestDatum() As Variant
    Dim Actual As String
    ReDim TestDatum(0 To 1)

    TestDatum(0) = Array("Sub-Saharan Africa", "Chad", "Office Supplies", "Online", "L", CDate("1/27/2011"), "292494523", CDate("2/12/2011"), 4484, 651.21, 524.96, 2920025.64, 2353920.64, 566105#)
    TestDatum(1) = Array("Europe", "Latvia", "Beverages", "Online", "C", CDate("12/28/2015"), "361825549", CDate("1/23/2016"), 1075, 47.45, 31.79, 51008.75, 34174.25, 16834.5)
    
    'Act:
    SUT.Items = TestDatum
    Actual = SUT.ToCSVString(DateFormat:="m/dd/yyyy", NumberFormat:="#,##0.00")
       
    ' TestUtils.PrintExpectedActualStringsToConsole expected, actual
       
    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BetterArray_ToCSVString")
Private Sub ToCSVString_Simple2RowNoHeadersEscapeCommasQuotesAndCRLF_ValidStringReturned()
    On Error GoTo TestFail

    'Arrange:
    Const Expected As String = _
        """Sub-Saharan, Africa"",Chad,""Office" & vbCrLf & "Supplies"",Online,L,1/27/2011,292494523,2/12/2011,4484,651.21,524.96,2920025.64,2353920.64,566105.00" & vbCrLf & _
        "Europe,Latvia,Bever""""ages,Online,C,12/28/2015,361825549,1/23/2016,1075,47.45,31.79,51008.75,34174.25,16834.50"

        
    Dim TestDatum() As Variant
    Dim Actual As String
    ReDim TestDatum(0 To 1)
    
    TestDatum(0) = Array("Sub-Saharan, Africa", "Chad", "Office" & vbCrLf & "Supplies", "Online", "L", "1/27/2011", "292494523", "2/12/2011", "4484", "651.21", "524.96", "2920025.64", "2353920.64", "566105.00")
    TestDatum(1) = Array("Europe", "Latvia", "Bever""ages", "Online", "C", "12/28/2015", "361825549", "1/23/2016", "1075", "47.45", "31.79", 51008.75, "34174.25", "16834.50")

    'Act:
    SUT.Items = TestDatum
    Actual = SUT.ToCSVString()
    
    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


