Attribute VB_Name = "TestModule_ExcelProvider"
Option Explicit
Option Private Module

'@TestModule
'@Folder("VBABetterArray.Tests.Dependencies.ExcelProvider.Tests")

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
Private SUT As ExcelProvider

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
    Set SUT = New ExcelProvider
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    Set SUT = Nothing
End Sub

'@TestMethod("ExcelProvider_Constructor")
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

'@TestMethod("ExcelProvider_ExcelApplication")
Private Sub ExcelApplication_ReturnsExcelInstance_InstanceIsCorrectType()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As String
    Dim Actual As String
    Expected = "Microsoft Excel"
    
    'Act:
    Actual = SUT.ExcelApplication.Name
    
    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ExcelProvider_CurrentWorkbook")
Private Sub CurrentWorkbook_ReturnsWorkbook_CurrentWorkbookNotNothing()
    On Error GoTo TestFail
    
    'Arrange:
    Dim returned As Object
    Dim Expected As String
    Dim Actual As String
    
    Expected = "Workbook"
    
    'Act:
    Set returned = SUT.CurrentWorkbook
    Actual = TypeName(returned)
    
    'Assert:
    Assert.AreEqual Expected, Actual, "Actual <> expected"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ExcelProvider_CurrentWorksheet")
Private Sub CurrentWorksheet_ReturnsWorksheet_ReturnsTypeWorksheet()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As String
    Dim Actual As String
    Dim returned As Object
    
    Expected = "Worksheet"
    
    'Act:
    Set returned = SUT.CurrentWorksheet
    Actual = TypeName(returned)
    
    'Assert:
    Assert.AreEqual Expected, Actual, "actual has incorrect type"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ExcelProvider_CurrentWorksheet")
Private Sub CurrentWorksheet_ReturnsWorksheet_WorksheetIsChildOfCurrentWorkbook()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As Object
    Dim Actual As Object
    
    'Act:
    Set Expected = SUT.CurrentWorkbook
    Set Actual = SUT.CurrentWorksheet
        
    'Assert:
    Assert.AreSame Expected, Actual.Parent, "actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ExcelProvider_CurrentWorksheet")
Private Sub CurrentWorksheet_CanSetRangeValue_ReturnsCorrectValue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim address As String
    Dim Expected As String
    Dim Actual As String
    
    address = "A1"
    Expected = "Hello World"
    
    'Act:
    SUT.CurrentWorksheet.Range(address) = Expected
    Actual = SUT.CurrentWorksheet.Range(address)
    
    'Assert:
    Assert.AreEqual Actual, Expected, "actual <> expected"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


