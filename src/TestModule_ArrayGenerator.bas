Attribute VB_Name = "TestModule_ArrayGenerator"
Option Explicit
Option Private Module

'@TestModule
'@Folder("VBABetterArray.Tests.Dependencies.ArrayGenerator.Tests")

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
Private SUT As ArrayGenerator

Private Const TEST_ARRAY_LENGTH As Long = 10

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
    Set SUT = New ArrayGenerator
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    Set SUT = Nothing
End Sub

'@TestMethod("ArrayGeneratorConstructor")
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

'@TestMethod("BooleanArrays")
Private Sub GetArray_BooleansOneDimension_ValuesAreBoolean()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    Dim element As Variant
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_BOOLEAN, AG_ArrayTypes.AG_ONEDIMENSION)
    testResult = True
    For Each element In returnedArray
        If TypeName(element) <> "Boolean" Then testResult = False
    Next
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BooleanArrays")
Private Sub GetArray_BooleansMultiDimension_ValuesAreBoolean()
    On Error GoTo TestFail

    'Arrange:
    Dim i As Long
    Dim j As Long
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_BOOLEAN, AG_ArrayTypes.AG_MULTIDIMENSION)
    testResult = True
    For i = LBound(returnedArray) To UBound(returnedArray)
        For j = LBound(returnedArray, 2) To UBound(returnedArray, 2)
            If TypeName(returnedArray(i, j)) <> "Boolean" Then testResult = False
        Next
    Next
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BooleanArrays")
Private Sub GetArray_BooleansJagged_ValuesAreBoolean()
    On Error GoTo TestFail

    'Arrange:
    Dim i As Long
    Dim j As Long
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_BOOLEAN, AG_ArrayTypes.AG_JAGGED)
    testResult = True
    For i = LBound(returnedArray) To UBound(returnedArray)
        For j = LBound(returnedArray(i)) To UBound(returnedArray(i))
            If TypeName(returnedArray(i)(j)) <> "Boolean" Then testResult = False
        Next
    Next
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("ByteArrays")
Private Sub GetArray_BytesOneDimension_ValuesAreBytes()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    Dim element As Variant
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_BYTE, AG_ArrayTypes.AG_ONEDIMENSION)
    testResult = True
    For Each element In returnedArray
        If TypeName(element) <> "Byte" Then testResult = False
    Next
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("ByteArrays")
Private Sub GetArray_BytesMultiDimension_ValuesAreBytes()
    On Error GoTo TestFail

    'Arrange:
    Dim i As Long
    Dim j As Long

    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_BYTE, AG_ArrayTypes.AG_MULTIDIMENSION)
    testResult = True
    For i = LBound(returnedArray) To UBound(returnedArray)
        For j = LBound(returnedArray, 2) To UBound(returnedArray, 2)
            If TypeName(returnedArray(i, j)) <> "Byte" Then testResult = False
        Next
    Next
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("ByteArrays")
Private Sub GetArray_BytesJagged_ValuesAreBytes()
    On Error GoTo TestFail

    'Arrange:
    Dim i As Long
    Dim j As Long
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_BYTE, AG_ArrayTypes.AG_JAGGED)
    testResult = True
    For i = LBound(returnedArray) To UBound(returnedArray)
        For j = LBound(returnedArray(i)) To UBound(returnedArray(i))
            If TypeName(returnedArray(i)(j)) <> "Byte" Then testResult = False
        Next
    Next
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("DoublesArrays")
Private Sub GetArray_DoublesOneDimension_ValuesAreDoubles()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    Dim element As Variant
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_DOUBLE, AG_ArrayTypes.AG_ONEDIMENSION)
    testResult = True
    For Each element In returnedArray
        If TypeName(element) <> "Double" Then testResult = False
    Next
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("DoublesArrays")
Private Sub GetArray_DoublesMultiDimension_ValuesAreDoubles()
    On Error GoTo TestFail

    'Arrange:
    Dim i As Long
    Dim j As Long
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_DOUBLE, AG_ArrayTypes.AG_MULTIDIMENSION)
    testResult = True
    For i = LBound(returnedArray) To UBound(returnedArray)
        For j = LBound(returnedArray, 2) To UBound(returnedArray, 2)
            If TypeName(returnedArray(i, j)) <> "Double" Then testResult = False
        Next
    Next
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("DoublesArrays")
Private Sub GetArray_DoublesJagged_ValuesAreDoubles()
    On Error GoTo TestFail

    'Arrange:
    Dim i As Long
    Dim j As Long
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_DOUBLE, AG_ArrayTypes.AG_JAGGED)
    testResult = True
    For i = LBound(returnedArray) To UBound(returnedArray)
        For j = LBound(returnedArray(i)) To UBound(returnedArray(i))
            If TypeName(returnedArray(i)(j)) <> "Double" Then testResult = False
        Next
    Next
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("LongsArrays")
Private Sub GetArray_LongsOneDimension_ValuesAreLongs()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    Dim element As Variant
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_LONG, AG_ArrayTypes.AG_ONEDIMENSION)
    testResult = True
    For Each element In returnedArray
        If TypeName(element) <> "Long" Then testResult = False
    Next
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("LongsArrays")
Private Sub GetArray_LongsMultiDimension_ValuesAreLongs()
    On Error GoTo TestFail

    'Arrange:
    Dim i As Long
    Dim j As Long

    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_LONG, AG_ArrayTypes.AG_MULTIDIMENSION)
    testResult = True
    For i = LBound(returnedArray) To UBound(returnedArray)
        For j = LBound(returnedArray, 2) To UBound(returnedArray, 2)
            If TypeName(returnedArray(i, j)) <> "Long" Then testResult = False
        Next
    Next
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("LongsArrays")
Private Sub GetArray_LongsJagged_ValuesAreLongs()
    On Error GoTo TestFail

    'Arrange:
    Dim i As Long
    Dim j As Long
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_LONG, AG_ArrayTypes.AG_JAGGED)
    testResult = True
    For i = LBound(returnedArray) To UBound(returnedArray)
        For j = LBound(returnedArray(i)) To UBound(returnedArray(i))
            If TypeName(returnedArray(i)(j)) <> "Long" Then testResult = False
        Next
    Next
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'Objects

'@TestMethod("ObjectsArrays")
Private Sub GetArray_ObjectsOneDimension_ValuesAreObjects()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    Dim element As Variant
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_OBJECT, AG_ArrayTypes.AG_ONEDIMENSION)
    testResult = True
    For Each element In returnedArray
        If Not IsObject(element) Then testResult = False
    Next
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("ObjectsArrays")
Private Sub GetArray_ObjectsMultiDimension_ValuesAreObjects()
    On Error GoTo TestFail

    'Arrange:
    Dim i As Long
    Dim j As Long
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_OBJECT, AG_ArrayTypes.AG_MULTIDIMENSION)
    testResult = True
    For i = LBound(returnedArray) To UBound(returnedArray)
        For j = LBound(returnedArray, 2) To UBound(returnedArray, 2)
            If Not IsObject(returnedArray(i, j)) Then testResult = False
        Next
    Next
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("ObjectsArrays")
Private Sub GetArray_ObjectsJagged_ValuesAreObjects()
    On Error GoTo TestFail

    'Arrange:
    Dim i As Long
    Dim j As Long
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_OBJECT, AG_ArrayTypes.AG_JAGGED)
    testResult = True
    For i = LBound(returnedArray) To UBound(returnedArray)
        For j = LBound(returnedArray(i)) To UBound(returnedArray(i))
            If Not IsObject(returnedArray(i)(j)) Then testResult = False
        Next
    Next
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'Strings


'@TestMethod("StringsArrays")
Private Sub GetArray_StringsOneDimension_ValuesAreStrings()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    Dim element As Variant
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_STRING, AG_ArrayTypes.AG_ONEDIMENSION)
    testResult = True
    For Each element In returnedArray
        If TypeName(element) <> "String" Then testResult = False
    Next
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("StringsArrays")
Private Sub GetArray_StringsMultiDimension_ValuesAreStrings()
    On Error GoTo TestFail

    'Arrange:
    Dim i As Long
    Dim j As Long
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_STRING, AG_ArrayTypes.AG_MULTIDIMENSION)
    testResult = True
    For i = LBound(returnedArray) To UBound(returnedArray)
        For j = LBound(returnedArray, 2) To UBound(returnedArray, 2)
            If TypeName(returnedArray(i, j)) <> "String" Then testResult = False
        Next
    Next
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("StringsArrays")
Private Sub GetArray_StringsJagged_ValuesAreStrings()
    On Error GoTo TestFail

    'Arrange:
    Dim i As Long
    Dim j As Long
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_STRING, AG_ArrayTypes.AG_JAGGED)
    testResult = True
    For i = LBound(returnedArray) To UBound(returnedArray)
        For j = LBound(returnedArray(i)) To UBound(returnedArray(i))
            If TypeName(returnedArray(i)(j)) <> "String" Then testResult = False
        Next
    Next
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'Variants

'@TestMethod("VariantsArrays")
Private Sub GetArray_VariantsOneDimension_ValueTypesVary()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    Dim firstType As String
    Dim element As Variant
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_VARIANT, AG_ArrayTypes.AG_ONEDIMENSION)
    firstType = TypeName(returnedArray(0))
    For Each element In returnedArray
        If TypeName(element) <> firstType Then
            testResult = True
            Exit For
        End If
    Next
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("VariantsArrays")
Private Sub GetArray_VariantsMultiDimension_ValueTypesVary()
    On Error GoTo TestFail

    'Arrange:
    Dim i As Long
    Dim j As Long
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    Dim firstType As String
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_VARIANT, AG_ArrayTypes.AG_MULTIDIMENSION)
    firstType = TypeName(returnedArray(0, 0))
    For i = LBound(returnedArray) To UBound(returnedArray)
        For j = LBound(returnedArray, 2) To UBound(returnedArray, 2)
            If TypeName(returnedArray(i, j)) <> firstType Then
                testResult = True
                Exit For
            End If
        Next
    Next
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("VariantsArrays")
Private Sub GetArray_VariantsJagged_ValueTypesVary()
    On Error GoTo TestFail

    'Arrange:
    Dim i As Long
    Dim j As Long
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    Dim firstType As String
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_VARIANT, AG_ArrayTypes.AG_JAGGED)
    firstType = TypeName(returnedArray(0)(0))
    For i = LBound(returnedArray) To UBound(returnedArray)
        For j = LBound(returnedArray(i)) To UBound(returnedArray(i))
            If TypeName(returnedArray(i)(j)) <> firstType Then
                testResult = True
                Exit For
            End If
        Next
    Next
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("BooleanArrays")
Private Sub GetArray_BooleanOneDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_BOOLEAN, AG_ArrayTypes.AG_ONEDIMENSION)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BooleanArrays")
Private Sub GetArray_BooleanMultiDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_BOOLEAN, AG_ArrayTypes.AG_MULTIDIMENSION)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BooleanArrays")
Private Sub GetArray_BooleanJagged_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_BOOLEAN, AG_ArrayTypes.AG_JAGGED)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("ByteArrays")
Private Sub GetArray_ByteOneDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_BYTE, AG_ArrayTypes.AG_ONEDIMENSION)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("ByteArrays")
Private Sub GetArray_ByteMultiDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_BYTE, AG_ArrayTypes.AG_MULTIDIMENSION)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("ByteArrays")
Private Sub GetArray_ByteJagged_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_BYTE, AG_ArrayTypes.AG_JAGGED)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("DoubleArrays")
Private Sub GetArray_DoubleOneDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_DOUBLE, AG_ArrayTypes.AG_ONEDIMENSION)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("DoubleArrays")
Private Sub GetArray_DoubleMultiDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_DOUBLE, AG_ArrayTypes.AG_MULTIDIMENSION)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("DoubleArrays")
Private Sub GetArray_DoubleJagged_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_DOUBLE, AG_ArrayTypes.AG_JAGGED)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("LongArrays")
Private Sub GetArray_LongOneDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_LONG, AG_ArrayTypes.AG_ONEDIMENSION)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("LongArrays")
Private Sub GetArray_LongMultiDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_LONG, AG_ArrayTypes.AG_MULTIDIMENSION)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("LongArrays")
Private Sub GetArray_LongJagged_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_LONG, AG_ArrayTypes.AG_JAGGED)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("ObjectArrays")
Private Sub GetArray_ObjectOneDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_OBJECT, AG_ArrayTypes.AG_ONEDIMENSION)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("ObjectArrays")
Private Sub GetArray_ObjectMultiDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_OBJECT, AG_ArrayTypes.AG_MULTIDIMENSION)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("ObjectArrays")
Private Sub GetArray_ObjectJagged_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_OBJECT, AG_ArrayTypes.AG_JAGGED)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("StringArrays")
Private Sub GetArray_StringOneDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_STRING, AG_ArrayTypes.AG_ONEDIMENSION)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("StringArrays")
Private Sub GetArray_StringMultiDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_STRING, AG_ArrayTypes.AG_MULTIDIMENSION)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("StringArrays")
Private Sub GetArray_StringJagged_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_STRING, AG_ArrayTypes.AG_JAGGED)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("VariantArrays")
Private Sub GetArray_VariantOneDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_VARIANT, AG_ArrayTypes.AG_ONEDIMENSION)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("VariantArrays")
Private Sub GetArray_VariantMultiDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_VARIANT, AG_ArrayTypes.AG_MULTIDIMENSION)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("VariantArrays")
Private Sub GetArray_VariantJagged_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_VARIANT, AG_ArrayTypes.AG_JAGGED)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BooleanArrays")
Private Sub GetArray_BooleanOneDimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_BOOLEAN, AG_ArrayTypes.AG_ONEDIMENSION)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BooleanArrays")
Private Sub GetArray_BooleanMultidimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_BOOLEAN, AG_ArrayTypes.AG_MULTIDIMENSION)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("BooleanArrays")
Private Sub GetArray_BooleanJagged_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_BOOLEAN, AG_ArrayTypes.AG_JAGGED)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("ByteArrays")
Private Sub GetArray_ByteOneDimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_BYTE, AG_ArrayTypes.AG_ONEDIMENSION)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("ByteArrays")
Private Sub GetArray_ByteMultidimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_BYTE, AG_ArrayTypes.AG_MULTIDIMENSION)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("ByteArrays")
Private Sub GetArray_ByteJagged_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_BYTE, AG_ArrayTypes.AG_JAGGED)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("DoubleArrays")
Private Sub GetArray_DoubleOneDimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_DOUBLE, AG_ArrayTypes.AG_ONEDIMENSION)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("DoubleArrays")
Private Sub GetArray_DoubleMultidimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_DOUBLE, AG_ArrayTypes.AG_MULTIDIMENSION)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("DoubleArrays")
Private Sub GetArray_DoubleJagged_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_DOUBLE, AG_ArrayTypes.AG_JAGGED)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("LongArrays")
Private Sub GetArray_LongOneDimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_LONG, AG_ArrayTypes.AG_ONEDIMENSION)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("LongArrays")
Private Sub GetArray_LongMultidimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_LONG, AG_ArrayTypes.AG_MULTIDIMENSION)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("LongArrays")
Private Sub GetArray_LongJagged_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_LONG, AG_ArrayTypes.AG_JAGGED)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("ObjectArrays")
Private Sub GetArray_ObjectOneDimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_OBJECT, AG_ArrayTypes.AG_ONEDIMENSION)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("ObjectArrays")
Private Sub GetArray_ObjectMultidimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_OBJECT, AG_ArrayTypes.AG_MULTIDIMENSION)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("ObjectArrays")
Private Sub GetArray_ObjectJagged_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_OBJECT, AG_ArrayTypes.AG_JAGGED)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("StringArrays")
Private Sub GetArray_StringOneDimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_STRING, AG_ArrayTypes.AG_ONEDIMENSION)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("StringArrays")
Private Sub GetArray_StringMultidimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_STRING, AG_ArrayTypes.AG_MULTIDIMENSION)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("StringArrays")
Private Sub GetArray_StringJagged_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_STRING, AG_ArrayTypes.AG_JAGGED)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("VariantArrays")
Private Sub GetArray_VariantOneDimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_VARIANT, AG_ArrayTypes.AG_ONEDIMENSION)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("VariantArrays")
Private Sub GetArray_VariantMultidimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_VARIANT, AG_ArrayTypes.AG_MULTIDIMENSION)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("VariantArrays")
Private Sub GetArray_VariantJagged_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim returnedArray() As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.GetArray(ValueTypes.AG_VARIANT, AG_ArrayTypes.AG_JAGGED)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub



'@TestMethod("ArrayGenerator_ConcatArraysOfSameStructure")
Private Sub ConcatArraysOfSameStructure_TwoMultiDimArrays_ConcatSuccess()
    On Error GoTo TestFail

    'Arrange:
    Dim firstArray(1, 1) As Variant
    Dim secondArray(1, 1) As Variant
    Dim expected(3, 1) As Variant
    Dim actual() As Variant
    
    firstArray(0, 0) = 0
    firstArray(0, 1) = "A"
    firstArray(1, 0) = 1
    firstArray(1, 1) = "B"
    secondArray(0, 0) = 2
    secondArray(0, 1) = "C"
    secondArray(1, 0) = 3
    secondArray(1, 1) = "D"
    
    expected(0, 0) = 0
    expected(0, 1) = "A"
    expected(1, 0) = 1
    expected(1, 1) = "B"
    expected(2, 0) = 2
    expected(2, 1) = "C"
    expected(3, 0) = 3
    expected(3, 1) = "D"
    
    'Act:
    actual = SUT.ConcatArraysOfSameStructure(AG_MULTIDIMENSION, firstArray, secondArray)
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("ArrayGenerator_ConcatTwoDimensionArrays")
Private Sub ConcatArraysOfSameStructures_ThreeMultiDimArrays_ConcatSuccess()
    On Error GoTo TestFail

    'Arrange:
    Dim firstArray(1, 1) As Variant
    Dim secondArray(1, 1) As Variant
    Dim thirdArray(1, 1) As Variant
    Dim expected(5, 1) As Variant
    Dim actual() As Variant
    
    firstArray(0, 0) = 0
    firstArray(0, 1) = "A"
    firstArray(1, 0) = 1
    firstArray(1, 1) = "B"
    secondArray(0, 0) = 2
    secondArray(0, 1) = "C"
    secondArray(1, 0) = 3
    secondArray(1, 1) = "D"
    thirdArray(0, 0) = 4
    thirdArray(0, 1) = "E"
    thirdArray(1, 0) = 5
    thirdArray(1, 1) = "F"
    
    expected(0, 0) = 0
    expected(0, 1) = "A"
    expected(1, 0) = 1
    expected(1, 1) = "B"
    expected(2, 0) = 2
    expected(2, 1) = "C"
    expected(3, 0) = 3
    expected(3, 1) = "D"
    expected(4, 0) = 4
    expected(4, 1) = "E"
    expected(5, 0) = 5
    expected(5, 1) = "F"
    
    'Act:
    actual = SUT.ConcatArraysOfSameStructure(AG_MULTIDIMENSION, firstArray, secondArray, thirdArray)
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("ArrayGenerator_ConcatArraysOfSameStructure")
Private Sub ConcatArraysOfSameStructure_TwoOneDimArrays_ConcatSuccess()
    On Error GoTo TestFail

    'Arrange:
    Dim firstArray(1) As Variant
    Dim secondArray(1) As Variant
    Dim expected(3) As Variant
    Dim actual() As Variant
    
    firstArray(0) = 0
    firstArray(1) = 1
    secondArray(0) = 2
    secondArray(1) = 3
    
    expected(0) = 0
    expected(1) = 1
    expected(2) = 2
    expected(3) = 3
    
    'Act:
    actual = SUT.ConcatArraysOfSameStructure(AG_ONEDIMENSION, firstArray, secondArray)
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("ArrayGenerator_ConcatTwoDimensionArrays")
Private Sub ConcatArraysOfSameStructures_ThreeOneDimArrays_ConcatSuccess()
    On Error GoTo TestFail

    'Arrange:
    Dim firstArray(1) As Variant
    Dim secondArray(1) As Variant
    Dim thirdArray(1) As Variant
    Dim expected(5) As Variant
    Dim actual() As Variant
    
    firstArray(0) = 0
    firstArray(1) = 1
    secondArray(0) = 2
    secondArray(1) = 3
    thirdArray(0) = 4
    thirdArray(1) = 5
    
    expected(0) = 0
    expected(1) = 1
    expected(2) = 2
    expected(3) = 3
    expected(4) = 4
    expected(5) = 5
    
    'Act:
    actual = SUT.ConcatArraysOfSameStructure(AG_ONEDIMENSION, firstArray, secondArray, thirdArray)
    
    'Assert:
    Assert.SequenceEquals expected, actual, "Actual <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("ArrayGenerator_ConcatArraysOfSameStructure")
Private Sub ConcatArraysOfSameStructure_TwoJaggedArrays_ConcatSuccess()
    On Error GoTo TestFail

    'Arrange:
    Dim firstArray() As Variant
    Dim secondArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    Dim i As Long
    
    firstArray = Array(Array(0, 1, 2), Array(3, 4, 5))
    secondArray = Array(Array(6, 7, 8), Array(9, 10, 11))
    expected = Array(Array(0, 1, 2), Array(3, 4, 5), Array(6, 7, 8), Array(9, 10, 11))
    
    'Act:
    actual = SUT.ConcatArraysOfSameStructure(AG_JAGGED, firstArray, secondArray)
    
    'Assert:
    For i = LBound(expected) To UBound(expected)
        Assert.SequenceEquals expected(i), actual(i), "Actual <> Expected"
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("ArrayGenerator_ConcatTwoDimensionArrays")
Private Sub ConcatArraysOfSameStructures_ThreeJaggedArrays_ConcatSuccess()
    On Error GoTo TestFail

    'Arrange:
    Dim firstArray() As Variant
    Dim secondArray() As Variant
    Dim thirdArray() As Variant
    Dim expected() As Variant
    Dim actual() As Variant
    Dim i As Long
    
    firstArray = Array(Array(0, 1, 2), Array(3, 4, 5))
    secondArray = Array(Array(6, 7, 8), Array(9, 10, 11))
    thirdArray = Array(Array(12, 13, 14), Array(15, 16, 17))
    expected = Array(Array(0, 1, 2), Array(3, 4, 5), Array(6, 7, 8), _
        Array(9, 10, 11), Array(12, 13, 14), Array(15, 16, 17))
    
    'Act:
    actual = SUT.ConcatArraysOfSameStructure(AG_JAGGED, firstArray, secondArray, thirdArray)
    
    'Assert:
    For i = LBound(expected) To UBound(expected)
        Assert.SequenceEquals expected(i), actual(i), "Actual <> Expected"
    Next
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("ArrayGenerator_GetArrayLength")
Private Sub GetArrayLength_OneDimArray_ReturnsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim testArray() As Variant
    Dim expected As Long
    Dim actual As Long
    
    expected = TEST_ARRAY_LENGTH
    
    'Act:
    testArray = SUT.GetArray(Length:=TEST_ARRAY_LENGTH)
    actual = SUT.GetArrayLength(testArray)
    
    'Assert:
    Assert.AreEqual expected, actual, "Actual <> Expected"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub
