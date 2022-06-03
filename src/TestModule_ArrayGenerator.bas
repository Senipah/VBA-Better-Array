Attribute VB_Name = "TestModule_ArrayGenerator"
Option Explicit
Option Private Module

'@TestModule
'@Folder("VBABetterArray.Tests.Dependencies.ArrayGenerator.Tests")

'@IgnoreModule VariableNotUsed, AssignmentNotUsed, ProcedureNotUsed, LineLabelNotUsed, EmptyMethod

' Uncomment for late binding
Private Assert As Object
' Move to early bind
'Private Assert As AssertClass

' Uncomment for late binding
Private Fakes As Object
' Uncomment for early  binding
'Private Fakes As FakesProvider

' Module level declaration of system under test
Private SUT As ArrayGenerator

Private Const TEST_ARRAY_LENGTH As Long = 10

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    ' Uncomment for late binding
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    ' Uncomment for early  binding
    ' Set Assert = New AssertClass
    ' Set Fakes = New FakesProvider
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
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BooleanArrays")
Private Sub GetArray_BooleansOneDimension_ValuesAreBoolean()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    Dim Element As Variant
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_BOOLEAN, AG_ArrayTypes.AG_ONEDIMENSION)
    TestResult = True
    For Each Element In ReturnedArray
        If TypeName(Element) <> "Boolean" Then TestResult = False
    Next
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BooleanArrays")
Private Sub GetArray_BooleansMultiDimension_ValuesAreBoolean()
    On Error GoTo TestFail

    'Arrange:
    Dim i As Long
    Dim j As Long
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_BOOLEAN, AG_ArrayTypes.AG_MULTIDIMENSION)
    TestResult = True
    For i = LBound(ReturnedArray) To UBound(ReturnedArray)
        For j = LBound(ReturnedArray, 2) To UBound(ReturnedArray, 2)
            If TypeName(ReturnedArray(i, j)) <> "Boolean" Then TestResult = False
        Next
    Next
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BooleanArrays")
Private Sub GetArray_BooleansJagged_ValuesAreBoolean()
    On Error GoTo TestFail

    'Arrange:
    Dim i As Long
    Dim j As Long
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_BOOLEAN, AG_ArrayTypes.AG_JAGGED)
    TestResult = True
    For i = LBound(ReturnedArray) To UBound(ReturnedArray)
        For j = LBound(ReturnedArray(i)) To UBound(ReturnedArray(i))
            If TypeName(ReturnedArray(i)(j)) <> "Boolean" Then TestResult = False
        Next
    Next
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ByteArrays")
Private Sub GetArray_BytesOneDimension_ValuesAreBytes()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    Dim Element As Variant
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_BYTE, AG_ArrayTypes.AG_ONEDIMENSION)
    TestResult = True
    For Each Element In ReturnedArray
        If TypeName(Element) <> "Byte" Then TestResult = False
    Next
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ByteArrays")
Private Sub GetArray_BytesMultiDimension_ValuesAreBytes()
    On Error GoTo TestFail

    'Arrange:
    Dim i As Long
    Dim j As Long

    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_BYTE, AG_ArrayTypes.AG_MULTIDIMENSION)
    TestResult = True
    For i = LBound(ReturnedArray) To UBound(ReturnedArray)
        For j = LBound(ReturnedArray, 2) To UBound(ReturnedArray, 2)
            If TypeName(ReturnedArray(i, j)) <> "Byte" Then TestResult = False
        Next
    Next
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ByteArrays")
Private Sub GetArray_BytesJagged_ValuesAreBytes()
    On Error GoTo TestFail

    'Arrange:
    Dim i As Long
    Dim j As Long
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_BYTE, AG_ArrayTypes.AG_JAGGED)
    TestResult = True
    For i = LBound(ReturnedArray) To UBound(ReturnedArray)
        For j = LBound(ReturnedArray(i)) To UBound(ReturnedArray(i))
            If TypeName(ReturnedArray(i)(j)) <> "Byte" Then TestResult = False
        Next
    Next
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DoublesArrays")
Private Sub GetArray_DoublesOneDimension_ValuesAreDoubles()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    Dim Element As Variant
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_DOUBLE, AG_ArrayTypes.AG_ONEDIMENSION)
    TestResult = True
    For Each Element In ReturnedArray
        If TypeName(Element) <> "Double" Then TestResult = False
    Next
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DoublesArrays")
Private Sub GetArray_DoublesMultiDimension_ValuesAreDoubles()
    On Error GoTo TestFail

    'Arrange:
    Dim i As Long
    Dim j As Long
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_DOUBLE, AG_ArrayTypes.AG_MULTIDIMENSION)
    TestResult = True
    For i = LBound(ReturnedArray) To UBound(ReturnedArray)
        For j = LBound(ReturnedArray, 2) To UBound(ReturnedArray, 2)
            If TypeName(ReturnedArray(i, j)) <> "Double" Then TestResult = False
        Next
    Next
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DoublesArrays")
Private Sub GetArray_DoublesJagged_ValuesAreDoubles()
    On Error GoTo TestFail

    'Arrange:
    Dim i As Long
    Dim j As Long
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_DOUBLE, AG_ArrayTypes.AG_JAGGED)
    TestResult = True
    For i = LBound(ReturnedArray) To UBound(ReturnedArray)
        For j = LBound(ReturnedArray(i)) To UBound(ReturnedArray(i))
            If TypeName(ReturnedArray(i)(j)) <> "Double" Then TestResult = False
        Next
    Next
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("LongsArrays")
Private Sub GetArray_LongsOneDimension_ValuesAreLongs()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    Dim Element As Variant
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_LONG, AG_ArrayTypes.AG_ONEDIMENSION)
    TestResult = True
    For Each Element In ReturnedArray
        If TypeName(Element) <> "Long" Then TestResult = False
    Next
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("LongsArrays")
Private Sub GetArray_LongsMultiDimension_ValuesAreLongs()
    On Error GoTo TestFail

    'Arrange:
    Dim i As Long
    Dim j As Long

    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_LONG, AG_ArrayTypes.AG_MULTIDIMENSION)
    TestResult = True
    For i = LBound(ReturnedArray) To UBound(ReturnedArray)
        For j = LBound(ReturnedArray, 2) To UBound(ReturnedArray, 2)
            If TypeName(ReturnedArray(i, j)) <> "Long" Then TestResult = False
        Next
    Next
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("LongsArrays")
Private Sub GetArray_LongsJagged_ValuesAreLongs()
    On Error GoTo TestFail

    'Arrange:
    Dim i As Long
    Dim j As Long
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_LONG, AG_ArrayTypes.AG_JAGGED)
    TestResult = True
    For i = LBound(ReturnedArray) To UBound(ReturnedArray)
        For j = LBound(ReturnedArray(i)) To UBound(ReturnedArray(i))
            If TypeName(ReturnedArray(i)(j)) <> "Long" Then TestResult = False
        Next
    Next
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'Objects

'@TestMethod("ObjectsArrays")
Private Sub GetArray_ObjectsOneDimension_ValuesAreObjects()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    Dim Element As Variant
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_OBJECT, AG_ArrayTypes.AG_ONEDIMENSION)
    TestResult = True
    For Each Element In ReturnedArray
        If Not IsObject(Element) Then TestResult = False
    Next
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ObjectsArrays")
Private Sub GetArray_ObjectsMultiDimension_ValuesAreObjects()
    On Error GoTo TestFail

    'Arrange:
    Dim i As Long
    Dim j As Long
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_OBJECT, AG_ArrayTypes.AG_MULTIDIMENSION)
    TestResult = True
    For i = LBound(ReturnedArray) To UBound(ReturnedArray)
        For j = LBound(ReturnedArray, 2) To UBound(ReturnedArray, 2)
            If Not IsObject(ReturnedArray(i, j)) Then TestResult = False
        Next
    Next
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ObjectsArrays")
Private Sub GetArray_ObjectsJagged_ValuesAreObjects()
    On Error GoTo TestFail

    'Arrange:
    Dim i As Long
    Dim j As Long
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_OBJECT, AG_ArrayTypes.AG_JAGGED)
    TestResult = True
    For i = LBound(ReturnedArray) To UBound(ReturnedArray)
        For j = LBound(ReturnedArray(i)) To UBound(ReturnedArray(i))
            If Not IsObject(ReturnedArray(i)(j)) Then TestResult = False
        Next
    Next
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'Strings


'@TestMethod("StringsArrays")
Private Sub GetArray_StringsOneDimension_ValuesAreStrings()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    Dim Element As Variant
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_STRING, AG_ArrayTypes.AG_ONEDIMENSION)
    TestResult = True
    For Each Element In ReturnedArray
        If TypeName(Element) <> "String" Then TestResult = False
    Next
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("StringsArrays")
Private Sub GetArray_StringsMultiDimension_ValuesAreStrings()
    On Error GoTo TestFail

    'Arrange:
    Dim i As Long
    Dim j As Long
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_STRING, AG_ArrayTypes.AG_MULTIDIMENSION)
    TestResult = True
    For i = LBound(ReturnedArray) To UBound(ReturnedArray)
        For j = LBound(ReturnedArray, 2) To UBound(ReturnedArray, 2)
            If TypeName(ReturnedArray(i, j)) <> "String" Then TestResult = False
        Next
    Next
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("StringsArrays")
Private Sub GetArray_StringsJagged_ValuesAreStrings()
    On Error GoTo TestFail

    'Arrange:
    Dim i As Long
    Dim j As Long
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_STRING, AG_ArrayTypes.AG_JAGGED)
    TestResult = True
    For i = LBound(ReturnedArray) To UBound(ReturnedArray)
        For j = LBound(ReturnedArray(i)) To UBound(ReturnedArray(i))
            If TypeName(ReturnedArray(i)(j)) <> "String" Then TestResult = False
        Next
    Next
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'Variants

'@TestMethod("VariantsArrays")
Private Sub GetArray_VariantsOneDimension_ValueTypesVary()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    Dim firstType As String
    Dim Element As Variant
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_VARIANT, AG_ArrayTypes.AG_ONEDIMENSION)
    firstType = TypeName(ReturnedArray(0))
    For Each Element In ReturnedArray
        If TypeName(Element) <> firstType Then
            TestResult = True
            Exit For
        End If
    Next
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("VariantsArrays")
Private Sub GetArray_VariantsMultiDimension_ValueTypesVary()
    On Error GoTo TestFail

    'Arrange:
    Dim i As Long
    Dim j As Long
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    Dim firstType As String
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_VARIANT, AG_ArrayTypes.AG_MULTIDIMENSION)
    firstType = TypeName(ReturnedArray(0, 0))
    For i = LBound(ReturnedArray) To UBound(ReturnedArray)
        For j = LBound(ReturnedArray, 2) To UBound(ReturnedArray, 2)
            If TypeName(ReturnedArray(i, j)) <> firstType Then
                TestResult = True
                Exit For
            End If
        Next
    Next
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("VariantsArrays")
Private Sub GetArray_VariantsJagged_ValueTypesVary()
    On Error GoTo TestFail

    'Arrange:
    Dim i As Long
    Dim j As Long
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    Dim firstType As String
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_VARIANT, AG_ArrayTypes.AG_JAGGED)
    firstType = TypeName(ReturnedArray(0)(0))
    For i = LBound(ReturnedArray) To UBound(ReturnedArray)
        For j = LBound(ReturnedArray(i)) To UBound(ReturnedArray(i))
            If TypeName(ReturnedArray(i)(j)) <> firstType Then
                TestResult = True
                Exit For
            End If
        Next
    Next
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BooleanArrays")
Private Sub GetArray_BooleanOneDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_BOOLEAN, AG_ArrayTypes.AG_ONEDIMENSION)
    TestResult = IsArray(ReturnedArray)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BooleanArrays")
Private Sub GetArray_BooleanMultiDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_BOOLEAN, AG_ArrayTypes.AG_MULTIDIMENSION)
    TestResult = IsArray(ReturnedArray)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BooleanArrays")
Private Sub GetArray_BooleanJagged_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_BOOLEAN, AG_ArrayTypes.AG_JAGGED)
    TestResult = IsArray(ReturnedArray)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ByteArrays")
Private Sub GetArray_ByteOneDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_BYTE, AG_ArrayTypes.AG_ONEDIMENSION)
    TestResult = IsArray(ReturnedArray)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ByteArrays")
Private Sub GetArray_ByteMultiDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_BYTE, AG_ArrayTypes.AG_MULTIDIMENSION)
    TestResult = IsArray(ReturnedArray)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ByteArrays")
Private Sub GetArray_ByteJagged_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_BYTE, AG_ArrayTypes.AG_JAGGED)
    TestResult = IsArray(ReturnedArray)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DoubleArrays")
Private Sub GetArray_DoubleOneDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_DOUBLE, AG_ArrayTypes.AG_ONEDIMENSION)
    TestResult = IsArray(ReturnedArray)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DoubleArrays")
Private Sub GetArray_DoubleMultiDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_DOUBLE, AG_ArrayTypes.AG_MULTIDIMENSION)
    TestResult = IsArray(ReturnedArray)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DoubleArrays")
Private Sub GetArray_DoubleJagged_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_DOUBLE, AG_ArrayTypes.AG_JAGGED)
    TestResult = IsArray(ReturnedArray)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("LongArrays")
Private Sub GetArray_LongOneDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_LONG, AG_ArrayTypes.AG_ONEDIMENSION)
    TestResult = IsArray(ReturnedArray)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("LongArrays")
Private Sub GetArray_LongMultiDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_LONG, AG_ArrayTypes.AG_MULTIDIMENSION)
    TestResult = IsArray(ReturnedArray)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("LongArrays")
Private Sub GetArray_LongJagged_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_LONG, AG_ArrayTypes.AG_JAGGED)
    TestResult = IsArray(ReturnedArray)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ObjectArrays")
Private Sub GetArray_ObjectOneDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_OBJECT, AG_ArrayTypes.AG_ONEDIMENSION)
    TestResult = IsArray(ReturnedArray)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ObjectArrays")
Private Sub GetArray_ObjectMultiDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_OBJECT, AG_ArrayTypes.AG_MULTIDIMENSION)
    TestResult = IsArray(ReturnedArray)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ObjectArrays")
Private Sub GetArray_ObjectJagged_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_OBJECT, AG_ArrayTypes.AG_JAGGED)
    TestResult = IsArray(ReturnedArray)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("StringArrays")
Private Sub GetArray_StringOneDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_STRING, AG_ArrayTypes.AG_ONEDIMENSION)
    TestResult = IsArray(ReturnedArray)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("StringArrays")
Private Sub GetArray_StringMultiDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_STRING, AG_ArrayTypes.AG_MULTIDIMENSION)
    TestResult = IsArray(ReturnedArray)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("StringArrays")
Private Sub GetArray_StringJagged_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_STRING, AG_ArrayTypes.AG_JAGGED)
    TestResult = IsArray(ReturnedArray)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("VariantArrays")
Private Sub GetArray_VariantOneDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_VARIANT, AG_ArrayTypes.AG_ONEDIMENSION)
    TestResult = IsArray(ReturnedArray)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("VariantArrays")
Private Sub GetArray_VariantMultiDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_VARIANT, AG_ArrayTypes.AG_MULTIDIMENSION)
    TestResult = IsArray(ReturnedArray)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("VariantArrays")
Private Sub GetArray_VariantJagged_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_VARIANT, AG_ArrayTypes.AG_JAGGED)
    TestResult = IsArray(ReturnedArray)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BooleanArrays")
Private Sub GetArray_BooleanOneDimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_BOOLEAN, AG_ArrayTypes.AG_ONEDIMENSION)
    TestResult = ((UBound(ReturnedArray) - LBound(ReturnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BooleanArrays")
Private Sub GetArray_BooleanMultidimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_BOOLEAN, AG_ArrayTypes.AG_MULTIDIMENSION)
    TestResult = ((UBound(ReturnedArray) - LBound(ReturnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BooleanArrays")
Private Sub GetArray_BooleanJagged_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_BOOLEAN, AG_ArrayTypes.AG_JAGGED)
    TestResult = ((UBound(ReturnedArray) - LBound(ReturnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ByteArrays")
Private Sub GetArray_ByteOneDimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_BYTE, AG_ArrayTypes.AG_ONEDIMENSION)
    TestResult = ((UBound(ReturnedArray) - LBound(ReturnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ByteArrays")
Private Sub GetArray_ByteMultidimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_BYTE, AG_ArrayTypes.AG_MULTIDIMENSION)
    TestResult = ((UBound(ReturnedArray) - LBound(ReturnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ByteArrays")
Private Sub GetArray_ByteJagged_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_BYTE, AG_ArrayTypes.AG_JAGGED)
    TestResult = ((UBound(ReturnedArray) - LBound(ReturnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DoubleArrays")
Private Sub GetArray_DoubleOneDimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_DOUBLE, AG_ArrayTypes.AG_ONEDIMENSION)
    TestResult = ((UBound(ReturnedArray) - LBound(ReturnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DoubleArrays")
Private Sub GetArray_DoubleMultidimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_DOUBLE, AG_ArrayTypes.AG_MULTIDIMENSION)
    TestResult = ((UBound(ReturnedArray) - LBound(ReturnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DoubleArrays")
Private Sub GetArray_DoubleJagged_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_DOUBLE, AG_ArrayTypes.AG_JAGGED)
    TestResult = ((UBound(ReturnedArray) - LBound(ReturnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("LongArrays")
Private Sub GetArray_LongOneDimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_LONG, AG_ArrayTypes.AG_ONEDIMENSION)
    TestResult = ((UBound(ReturnedArray) - LBound(ReturnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("LongArrays")
Private Sub GetArray_LongMultidimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_LONG, AG_ArrayTypes.AG_MULTIDIMENSION)
    TestResult = ((UBound(ReturnedArray) - LBound(ReturnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("LongArrays")
Private Sub GetArray_LongJagged_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_LONG, AG_ArrayTypes.AG_JAGGED)
    TestResult = ((UBound(ReturnedArray) - LBound(ReturnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ObjectArrays")
Private Sub GetArray_ObjectOneDimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_OBJECT, AG_ArrayTypes.AG_ONEDIMENSION)
    TestResult = ((UBound(ReturnedArray) - LBound(ReturnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ObjectArrays")
Private Sub GetArray_ObjectMultidimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_OBJECT, AG_ArrayTypes.AG_MULTIDIMENSION)
    TestResult = ((UBound(ReturnedArray) - LBound(ReturnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ObjectArrays")
Private Sub GetArray_ObjectJagged_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_OBJECT, AG_ArrayTypes.AG_JAGGED)
    TestResult = ((UBound(ReturnedArray) - LBound(ReturnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("StringArrays")
Private Sub GetArray_StringOneDimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_STRING, AG_ArrayTypes.AG_ONEDIMENSION)
    TestResult = ((UBound(ReturnedArray) - LBound(ReturnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("StringArrays")
Private Sub GetArray_StringMultidimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_STRING, AG_ArrayTypes.AG_MULTIDIMENSION)
    TestResult = ((UBound(ReturnedArray) - LBound(ReturnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("StringArrays")
Private Sub GetArray_StringJagged_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_STRING, AG_ArrayTypes.AG_JAGGED)
    TestResult = ((UBound(ReturnedArray) - LBound(ReturnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("VariantArrays")
Private Sub GetArray_VariantOneDimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_VARIANT, AG_ArrayTypes.AG_ONEDIMENSION)
    TestResult = ((UBound(ReturnedArray) - LBound(ReturnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("VariantArrays")
Private Sub GetArray_VariantMultidimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_VARIANT, AG_ArrayTypes.AG_MULTIDIMENSION)
    TestResult = ((UBound(ReturnedArray) - LBound(ReturnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("VariantArrays")
Private Sub GetArray_VariantJagged_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim ReturnedArray() As Variant
    Dim TestResult As Boolean
    
    'Act:
    ReturnedArray = SUT.GetArray(ValueTypes.AG_VARIANT, AG_ArrayTypes.AG_JAGGED)
    TestResult = ((UBound(ReturnedArray) - LBound(ReturnedArray) + 1) = TEST_ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue TestResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub



'@TestMethod("ArrayGenerator_ConcatArraysOfSameStructure")
Private Sub ConcatArraysOfSameStructure_TwoMultiDimArrays_ConcatSuccess()
    On Error GoTo TestFail

    'Arrange:
    Dim FirstArray(1, 1) As Variant
    Dim SecondArray(1, 1) As Variant
    Dim Expected(3, 1) As Variant
    Dim Actual() As Variant
    
    FirstArray(0, 0) = 0
    FirstArray(0, 1) = "A"
    FirstArray(1, 0) = 1
    FirstArray(1, 1) = "B"
    SecondArray(0, 0) = 2
    SecondArray(0, 1) = "C"
    SecondArray(1, 0) = 3
    SecondArray(1, 1) = "D"
    
    Expected(0, 0) = 0
    Expected(0, 1) = "A"
    Expected(1, 0) = 1
    Expected(1, 1) = "B"
    Expected(2, 0) = 2
    Expected(2, 1) = "C"
    Expected(3, 0) = 3
    Expected(3, 1) = "D"
    
    'Act:
    Actual = SUT.ConcatArraysOfSameStructure(AG_MULTIDIMENSION, FirstArray, SecondArray)
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ArrayGenerator_ConcatTwoDimensionArrays")
Private Sub ConcatArraysOfSameStructures_ThreeMultiDimArrays_ConcatSuccess()
    On Error GoTo TestFail

    'Arrange:
    Dim FirstArray(1, 1) As Variant
    Dim SecondArray(1, 1) As Variant
    Dim thirdArray(1, 1) As Variant
    Dim Expected(5, 1) As Variant
    Dim Actual() As Variant
    
    FirstArray(0, 0) = 0
    FirstArray(0, 1) = "A"
    FirstArray(1, 0) = 1
    FirstArray(1, 1) = "B"
    SecondArray(0, 0) = 2
    SecondArray(0, 1) = "C"
    SecondArray(1, 0) = 3
    SecondArray(1, 1) = "D"
    thirdArray(0, 0) = 4
    thirdArray(0, 1) = "E"
    thirdArray(1, 0) = 5
    thirdArray(1, 1) = "F"
    
    Expected(0, 0) = 0
    Expected(0, 1) = "A"
    Expected(1, 0) = 1
    Expected(1, 1) = "B"
    Expected(2, 0) = 2
    Expected(2, 1) = "C"
    Expected(3, 0) = 3
    Expected(3, 1) = "D"
    Expected(4, 0) = 4
    Expected(4, 1) = "E"
    Expected(5, 0) = 5
    Expected(5, 1) = "F"
    
    'Act:
    Actual = SUT.ConcatArraysOfSameStructure(AG_MULTIDIMENSION, FirstArray, SecondArray, thirdArray)
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ArrayGenerator_ConcatArraysOfSameStructure")
Private Sub ConcatArraysOfSameStructure_TwoOneDimArrays_ConcatSuccess()
    On Error GoTo TestFail

    'Arrange:
    Dim FirstArray(1) As Variant
    Dim SecondArray(1) As Variant
    Dim Expected(3) As Variant
    Dim Actual() As Variant
    
    FirstArray(0) = 0
    FirstArray(1) = 1
    SecondArray(0) = 2
    SecondArray(1) = 3
    
    Expected(0) = 0
    Expected(1) = 1
    Expected(2) = 2
    Expected(3) = 3
    
    'Act:
    Actual = SUT.ConcatArraysOfSameStructure(AG_ONEDIMENSION, FirstArray, SecondArray)
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ArrayGenerator_ConcatTwoDimensionArrays")
Private Sub ConcatArraysOfSameStructures_ThreeOneDimArrays_ConcatSuccess()
    On Error GoTo TestFail

    'Arrange:
    Dim FirstArray(1) As Variant
    Dim SecondArray(1) As Variant
    Dim thirdArray(1) As Variant
    Dim Expected(5) As Variant
    Dim Actual() As Variant
    
    FirstArray(0) = 0
    FirstArray(1) = 1
    SecondArray(0) = 2
    SecondArray(1) = 3
    thirdArray(0) = 4
    thirdArray(1) = 5
    
    Expected(0) = 0
    Expected(1) = 1
    Expected(2) = 2
    Expected(3) = 3
    Expected(4) = 4
    Expected(5) = 5
    
    'Act:
    Actual = SUT.ConcatArraysOfSameStructure(AG_ONEDIMENSION, FirstArray, SecondArray, thirdArray)
    
    'Assert:
    Assert.SequenceEquals Expected, Actual, "Actual <> Expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ArrayGenerator_ConcatArraysOfSameStructure")
Private Sub ConcatArraysOfSameStructure_TwoJaggedArrays_ConcatSuccess()
    On Error GoTo TestFail

    'Arrange:
    Dim FirstArray() As Variant
    Dim SecondArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim i As Long
    
    FirstArray = Array(Array(0, 1, 2), Array(3, 4, 5))
    SecondArray = Array(Array(6, 7, 8), Array(9, 10, 11))
    Expected = Array(Array(0, 1, 2), Array(3, 4, 5), Array(6, 7, 8), Array(9, 10, 11))
    
    'Act:
    Actual = SUT.ConcatArraysOfSameStructure(AG_JAGGED, FirstArray, SecondArray)
    
    'Assert:
    For i = LBound(Expected) To UBound(Expected)
        Assert.SequenceEquals Expected(i), Actual(i), "Actual <> Expected"
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ArrayGenerator_ConcatTwoDimensionArrays")
Private Sub ConcatArraysOfSameStructures_ThreeJaggedArrays_ConcatSuccess()
    On Error GoTo TestFail

    'Arrange:
    Dim FirstArray() As Variant
    Dim SecondArray() As Variant
    Dim thirdArray() As Variant
    Dim Expected() As Variant
    Dim Actual() As Variant
    Dim i As Long
    
    FirstArray = Array(Array(0, 1, 2), Array(3, 4, 5))
    SecondArray = Array(Array(6, 7, 8), Array(9, 10, 11))
    thirdArray = Array(Array(12, 13, 14), Array(15, 16, 17))
    Expected = Array(Array(0, 1, 2), Array(3, 4, 5), Array(6, 7, 8), _
        Array(9, 10, 11), Array(12, 13, 14), Array(15, 16, 17))
    
    'Act:
    Actual = SUT.ConcatArraysOfSameStructure(AG_JAGGED, FirstArray, SecondArray, thirdArray)
    
    'Assert:
    For i = LBound(Expected) To UBound(Expected)
        Assert.SequenceEquals Expected(i), Actual(i), "Actual <> Expected"
    Next
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ArrayGenerator_GetArrayLength")
Private Sub GetArrayLength_OneDimArray_ReturnsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim TestArray() As Variant
    Dim Expected As Long
    Dim Actual As Long
    
    Expected = TEST_ARRAY_LENGTH
    
    'Act:
    TestArray = SUT.GetArray(Length:=TEST_ARRAY_LENGTH)
    Actual = SUT.GetArrayLength(TestArray)
    
    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> Expected"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

