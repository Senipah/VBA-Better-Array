Attribute VB_Name = "TestModule_ArrayGenerator"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Dependencies.ArrayGenerator.Tests")

'@IgnoreModule ProcedureNotUsed
'@IgnoreModule LineLabelNotUsed
'@IgnoreModule EmptyMethod

Private Assert As Object
'@Ignore VariableNotUsed
Private Fakes As Object

Private Const ARRAY_LENGTH As Long = 10

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

'@TestMethod("ArrayGeneratorConstructor")
Private Sub CanInstantiate_SUTNotNothing()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SUT As ArrayGenerator
    'Act:
    Set SUT = New ArrayGenerator
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
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    Dim Element As Variant
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.BooleanVals, ArrayTypes.OneDimension)
    testResult = True
    For Each Element In returnedArray
        If TypeName(Element) <> "Boolean" Then testResult = False
    Next
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BooleanArrays")
Private Sub GetArray_BooleansMultiDimension_ValuesAreBoolean()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim i As Long
Dim j As Long

    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.BooleanVals, ArrayTypes.MultiDimension)
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
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BooleanArrays")
Private Sub GetArray_BooleansJagged_ValuesAreBoolean()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim i As Long
Dim j As Long

    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.BooleanVals, ArrayTypes.Jagged)
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
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ByteArrays")
Private Sub GetArray_BytesOneDimension_ValuesAreBytes()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    Dim Element As Variant
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.ByteVals, ArrayTypes.OneDimension)
    testResult = True
    For Each Element In returnedArray
        If TypeName(Element) <> "Byte" Then testResult = False
    Next
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ByteArrays")
Private Sub GetArray_BytesMultiDimension_ValuesAreBytes()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim i As Long
Dim j As Long

    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.ByteVals, ArrayTypes.MultiDimension)
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
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ByteArrays")
Private Sub GetArray_BytesJagged_ValuesAreBytes()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim i As Long
Dim j As Long

    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.ByteVals, ArrayTypes.Jagged)
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
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DoublesArrays")
Private Sub GetArray_DoublesOneDimension_ValuesAreDoubles()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    Dim Element As Variant
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.DoubleVals, ArrayTypes.OneDimension)
    testResult = True
    For Each Element In returnedArray
        If TypeName(Element) <> "Double" Then testResult = False
    Next
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DoublesArrays")
Private Sub GetArray_DoublesMultiDimension_ValuesAreDoubles()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim i As Long
Dim j As Long

    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.DoubleVals, ArrayTypes.MultiDimension)
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
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DoublesArrays")
Private Sub GetArray_DoublesJagged_ValuesAreDoubles()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim i As Long
Dim j As Long

    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.DoubleVals, ArrayTypes.Jagged)
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
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("LongsArrays")
Private Sub GetArray_LongsOneDimension_ValuesAreLongs()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    Dim Element As Variant
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.LongVals, ArrayTypes.OneDimension)
    testResult = True
    For Each Element In returnedArray
        If TypeName(Element) <> "Long" Then testResult = False
    Next
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("LongsArrays")
Private Sub GetArray_LongsMultiDimension_ValuesAreLongs()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim i As Long
Dim j As Long

    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.LongVals, ArrayTypes.MultiDimension)
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
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("LongsArrays")
Private Sub GetArray_LongsJagged_ValuesAreLongs()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim i As Long
Dim j As Long

    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.LongVals, ArrayTypes.Jagged)
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
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'Objects

'@TestMethod("ObjectsArrays")
Private Sub GetArray_ObjectsOneDimension_ValuesAreObjects()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    Dim Element As Variant
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.ObjectVals, ArrayTypes.OneDimension)
    testResult = True
    For Each Element In returnedArray
        If Not IsObject(Element) Then testResult = False
    Next
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ObjectsArrays")
Private Sub GetArray_ObjectsMultiDimension_ValuesAreObjects()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim i As Long
Dim j As Long

    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.ObjectVals, ArrayTypes.MultiDimension)
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
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ObjectsArrays")
Public Sub GetArray_ObjectsJagged_ValuesAreObjects()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim i As Long
Dim j As Long

    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.ObjectVals, ArrayTypes.Jagged)
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
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'Strings


'@TestMethod("StringsArrays")
Private Sub GetArray_StringsOneDimension_ValuesAreStrings()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    Dim Element As Variant
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.StringVals, ArrayTypes.OneDimension)
    testResult = True
    For Each Element In returnedArray
        If TypeName(Element) <> "String" Then testResult = False
    Next
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("StringsArrays")
Private Sub GetArray_StringsMultiDimension_ValuesAreStrings()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim i As Long
Dim j As Long

    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.StringVals, ArrayTypes.MultiDimension)
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
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("StringsArrays")
Private Sub GetArray_StringsJagged_ValuesAreStrings()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim i As Long
Dim j As Long

    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.StringVals, ArrayTypes.Jagged)
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
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'Variants

'@TestMethod("VariantsArrays")
Private Sub GetArray_VariantsOneDimension_ValueTypesVary()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    Dim firstType As String
    Dim Element As Variant
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.VariantVals, ArrayTypes.OneDimension)
    firstType = TypeName(returnedArray(0))
    For Each Element In returnedArray
        If TypeName(Element) <> firstType Then
            testResult = True
            Exit For
        End If
    Next
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("VariantsArrays")
Private Sub GetArray_VariantsMultiDimension_ValueTypesVary()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim i As Long
Dim j As Long

    Dim returnedArray As Variant
    Dim testResult As Boolean
    Dim firstType As String
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.VariantVals, ArrayTypes.MultiDimension)
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
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("VariantsArrays")
Private Sub GetArray_VariantsJagged_ValueTypesVary()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim i As Long
Dim j As Long

    Dim returnedArray As Variant
    Dim testResult As Boolean
    Dim firstType As String
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.VariantVals, ArrayTypes.Jagged)
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
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("BooleanArrays")
Private Sub GetArray_BooleanOneDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.BooleanVals, ArrayTypes.OneDimension)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BooleanArrays")
Private Sub GetArray_BooleanMultiDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.BooleanVals, ArrayTypes.MultiDimension)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BooleanArrays")
Private Sub GetArray_BooleanJagged_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.BooleanVals, ArrayTypes.Jagged)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ByteArrays")
Private Sub GetArray_ByteOneDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.ByteVals, ArrayTypes.OneDimension)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ByteArrays")
Private Sub GetArray_ByteMultiDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.ByteVals, ArrayTypes.MultiDimension)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ByteArrays")
Private Sub GetArray_ByteJagged_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.ByteVals, ArrayTypes.Jagged)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DoubleArrays")
Private Sub GetArray_DoubleOneDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.DoubleVals, ArrayTypes.OneDimension)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DoubleArrays")
Private Sub GetArray_DoubleMultiDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.DoubleVals, ArrayTypes.MultiDimension)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DoubleArrays")
Private Sub GetArray_DoubleJagged_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.DoubleVals, ArrayTypes.Jagged)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("LongArrays")
Private Sub GetArray_LongOneDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.LongVals, ArrayTypes.OneDimension)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("LongArrays")
Private Sub GetArray_LongMultiDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.LongVals, ArrayTypes.MultiDimension)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("LongArrays")
Private Sub GetArray_LongJagged_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.LongVals, ArrayTypes.Jagged)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ObjectArrays")
Private Sub GetArray_ObjectOneDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.ObjectVals, ArrayTypes.OneDimension)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ObjectArrays")
Private Sub GetArray_ObjectMultiDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.ObjectVals, ArrayTypes.MultiDimension)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ObjectArrays")
Private Sub GetArray_ObjectJagged_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.ObjectVals, ArrayTypes.Jagged)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("StringArrays")
Private Sub GetArray_StringOneDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.StringVals, ArrayTypes.OneDimension)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("StringArrays")
Private Sub GetArray_StringMultiDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.StringVals, ArrayTypes.MultiDimension)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("StringArrays")
Private Sub GetArray_StringJagged_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.StringVals, ArrayTypes.Jagged)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("VariantArrays")
Private Sub GetArray_VariantOneDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.VariantVals, ArrayTypes.OneDimension)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("VariantArrays")
Private Sub GetArray_VariantMultiDimension_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.VariantVals, ArrayTypes.MultiDimension)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("VariantArrays")
Private Sub GetArray_VariantJagged_ReturnsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.VariantVals, ArrayTypes.Jagged)
    testResult = IsArray(returnedArray)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'    Booleans
'    Bytes
'    Doubles
'    Longs
'    Objects
'    Strings
'    Variants
'
'    OneDimension
'    MultiDimension
'    jagged


Private Function getLength(ByRef arr() As Variant) As Long
    getLength = UBound(arr) - LBound(arr)
End Function

'@TestMethod("BooleanArrays")
Private Sub GetArray_BooleanOneDimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.BooleanVals, ArrayTypes.OneDimension)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BooleanArrays")
Private Sub GetArray_BooleanMultidimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.BooleanVals, ArrayTypes.MultiDimension)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("BooleanArrays")
Private Sub GetArray_BooleanJagged_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.BooleanVals, ArrayTypes.Jagged)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ByteArrays")
Private Sub GetArray_ByteOneDimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.ByteVals, ArrayTypes.OneDimension)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ByteArrays")
Private Sub GetArray_ByteMultidimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.ByteVals, ArrayTypes.MultiDimension)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ByteArrays")
Private Sub GetArray_ByteJagged_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.ByteVals, ArrayTypes.Jagged)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DoubleArrays")
Private Sub GetArray_DoubleOneDimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.DoubleVals, ArrayTypes.OneDimension)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DoubleArrays")
Private Sub GetArray_DoubleMultidimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.DoubleVals, ArrayTypes.MultiDimension)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("DoubleArrays")
Private Sub GetArray_DoubleJagged_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.DoubleVals, ArrayTypes.Jagged)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("LongArrays")
Private Sub GetArray_LongOneDimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.LongVals, ArrayTypes.OneDimension)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("LongArrays")
Private Sub GetArray_LongMultidimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.LongVals, ArrayTypes.MultiDimension)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("LongArrays")
Private Sub GetArray_LongJagged_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.LongVals, ArrayTypes.Jagged)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ObjectArrays")
Private Sub GetArray_ObjectOneDimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.ObjectVals, ArrayTypes.OneDimension)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ObjectArrays")
Private Sub GetArray_ObjectMultidimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.ObjectVals, ArrayTypes.MultiDimension)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ObjectArrays")
Private Sub GetArray_ObjectJagged_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.ObjectVals, ArrayTypes.Jagged)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("StringArrays")
Private Sub GetArray_StringOneDimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.StringVals, ArrayTypes.OneDimension)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("StringArrays")
Private Sub GetArray_StringMultidimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.StringVals, ArrayTypes.MultiDimension)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("StringArrays")
Private Sub GetArray_StringJagged_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.StringVals, ArrayTypes.Jagged)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("VariantArrays")
Private Sub GetArray_VariantOneDimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.VariantVals, ArrayTypes.OneDimension)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("VariantArrays")
Private Sub GetArray_VariantMultidimension_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.VariantVals, ArrayTypes.MultiDimension)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("VariantArrays")
Private Sub GetArray_VariantJagged_IsCorrectLength()
    On Error GoTo TestFail

    'Arrange:
    Dim SUT As ArrayGenerator
    Set SUT = New ArrayGenerator
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(ARRAY_LENGTH, ValueTypes.VariantVals, ArrayTypes.Jagged)
    testResult = ((UBound(returnedArray) - LBound(returnedArray) + 1) = ARRAY_LENGTH)
    
    'Assert:
    Assert.IsTrue testResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub



