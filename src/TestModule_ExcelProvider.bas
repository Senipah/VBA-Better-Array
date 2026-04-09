Attribute VB_Name = "TestModule_ExcelProvider"
Option Explicit
Option Private Module



Private Assert As AssertClass

' Module level declaration of system under test
Private SUT As ExcelProvider

Public Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New AssertClass
End Sub

Public Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
End Sub

Public Sub TestInitialize()
    'this method runs before every test in the module.
    Set SUT = New ExcelProvider
End Sub

Public Sub TestCleanup()
    'this method runs after every test in the module.
    Set SUT = Nothing
End Sub

Public Sub Constructor_CanInstantiate_SUTNotNothing()
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

Public Sub ExcelApplication_ReturnsExcelInstance_InstanceIsCorrectType()
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


Public Sub CurrentWorkbook_ReturnsWorkbook_CurrentWorkbookNotNothing()
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

Public Sub CurrentWorksheet_ReturnsWorksheet_ReturnsTypeWorksheet()
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

Public Sub CurrentWorksheet_ReturnsWorksheet_WorksheetIsChildOfCurrentWorkbook()
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

Public Sub CurrentWorksheet_CanSetRangeValue_ReturnsCorrectValue()
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

