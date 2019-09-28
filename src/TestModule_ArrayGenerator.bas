Attribute VB_Name = "TestModule_ArrayGenerator"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

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
Private Sub CanInstantiate_SUTNotNothing()                        'TODO Rename test
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
    Dim v As Variant
    
    'Act:
    returnedArray = SUT.getArray(10, Booleans, OneDimension)
    testResult = True
    For Each v In returnedArray
        If TypeName(v) <> "Boolean" Then testResult = False
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
    Dim i As Long, j As Long
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(10, Booleans, MultiDimension)
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
    Dim i As Long, j As Long
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(10, Booleans, jagged)
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
    Dim v As Variant
    
    'Act:
    returnedArray = SUT.getArray(10, Bytes, OneDimension)
    testResult = True
    For Each v In returnedArray
        If TypeName(v) <> "Byte" Then testResult = False
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
    Dim i As Long, j As Long
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(10, Bytes, MultiDimension)
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
    Dim i As Long, j As Long
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(10, Bytes, jagged)
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
    Dim v As Variant
    
    'Act:
    returnedArray = SUT.getArray(10, Doubles, OneDimension)
    testResult = True
    For Each v In returnedArray
        If TypeName(v) <> "Double" Then testResult = False
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
    Dim i As Long, j As Long
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(10, Doubles, MultiDimension)
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
    Dim i As Long, j As Long
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(10, Doubles, jagged)
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
    Dim v As Variant
    
    'Act:
    returnedArray = SUT.getArray(10, Longs, OneDimension)
    testResult = True
    For Each v In returnedArray
        If TypeName(v) <> "Long" Then testResult = False
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
    Dim i As Long, j As Long
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(10, Longs, MultiDimension)
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
    Dim i As Long, j As Long
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(10, Longs, jagged)
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
    Dim v As Variant
    
    'Act:
    returnedArray = SUT.getArray(10, Objects, OneDimension)
    testResult = True
    For Each v In returnedArray
        If Not IsObject(v) Then testResult = False
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
    Dim i As Long, j As Long
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(10, Objects, MultiDimension)
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
    Dim i As Long, j As Long
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(10, Objects, jagged)
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
    Dim v As Variant
    
    'Act:
    returnedArray = SUT.getArray(10, Strings, OneDimension)
    testResult = True
    For Each v In returnedArray
        If TypeName(v) <> "String" Then testResult = False
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
    Dim i As Long, j As Long
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(10, Strings, MultiDimension)
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
    Dim i As Long, j As Long
    Dim returnedArray As Variant
    Dim testResult As Boolean
    
    'Act:
    returnedArray = SUT.getArray(10, Strings, jagged)
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
    Dim v As Variant
    
    'Act:
    returnedArray = SUT.getArray(10, Variants, OneDimension)
    firstType = TypeName(returnedArray(0))
    For Each v In returnedArray
        If TypeName(v) <> firstType Then
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
    Dim i As Long, j As Long
    Dim returnedArray As Variant
    Dim testResult As Boolean
    Dim firstType As String
    
    'Act:
    returnedArray = SUT.getArray(10, Variants, MultiDimension)
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
    Dim i As Long, j As Long
    Dim returnedArray As Variant
    Dim testResult As Boolean
    Dim firstType As String
    
    'Act:
    returnedArray = SUT.getArray(10, Variants, jagged)
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
