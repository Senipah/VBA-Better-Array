VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "An RAII container providing access to an Excel application. Dependency for 'TestModule_BetterArray'"
Option Explicit

'@Folder("Tests.Dependencies.ExcelProvider")
'@ModuleDescription("An RAII container providing access to an Excel application. Dependency for 'TestModule_BetterArray'")

Private pExcelApplication As Object
Private pCurrentWorkbook As Object
Private pCurrentWorksheet As Object
Private pCurrentRange As Object
Private Const ErrCode As Long = vbObjectError + 513

'@Description("Constructor")
Private Sub Class_Initialize()
Attribute Class_Initialize.VB_Description = "Constructor"
    On Error GoTo ErrHandler
    Set pExcelApplication = CreateObject("Excel.Application")
    Set pCurrentWorkbook = pExcelApplication.Workbooks.Add
    Set pCurrentWorksheet = pCurrentWorkbook.Worksheets.Add
    On Error GoTo 0
    Exit Sub
ErrHandler:
    On Error Resume Next
    pExcelApplication.Visible = True
    Err.Raise ErrCode, TypeName(Me), "An error ocurred during initialization."
    On Error GoTo 0
End Sub

'@Description("Destructor")
Private Sub Class_Terminate()
Attribute Class_Terminate.VB_Description = "Destructor"
    On Error GoTo ErrHandler
    If Not pCurrentWorkbook Is Nothing Then
        pCurrentWorkbook.Close SaveChanges:=False
    End If
    If Not pExcelApplication Is Nothing Then
        pExcelApplication.Quit
    End If
    If Not pExcelApplication Is Nothing Then
        Set pExcelApplication = Nothing
    End If
    On Error GoTo 0
    Exit Sub
ErrHandler:
    On Error Resume Next
    pExcelApplication.Visible = True
    Err.Raise ErrCode, TypeName(Me), "An error ocurred in the destructor." _
        & vbCrLf & "Ensure no orphan Excel application processes exist on system."
    On Error GoTo 0
End Sub

'@Description("Returns the acquired Excel Application object")
Public Property Get ExcelApplication() As Object
Attribute ExcelApplication.VB_Description = "Returns the acquired Excel Application object"
    Set ExcelApplication = pExcelApplication
End Property

'@Description("Returns the Excel workbook object")
Public Property Get CurrentWorkbook() As Object
Attribute CurrentWorkbook.VB_Description = "Returns the Excel workbook object"
    Set CurrentWorkbook = pCurrentWorkbook
End Property

'@Description("Returns the active Range object")
Public Property Get CurrentRange() As Object
Attribute CurrentRange.VB_Description = "Returns the active Range object"
    If pCurrentRange Is Nothing Then
        Set pCurrentRange = pCurrentWorksheet.Range("A1")
    End If
    Set CurrentRange = pCurrentRange
End Property

'@Description("Sets the active Range object")
Public Property Set CurrentRange(ByRef NewRange As Object)
Attribute CurrentRange.VB_Description = "Sets the active Range object"
    On Error GoTo ErrHandler
    If TypeName(NewRange) <> "Range" Then GoTo ErrHandler
    If Not NewRange.Parent Is pCurrentWorksheet Then GoTo ErrHandler
    Set pCurrentRange = NewRange
    On Error GoTo 0
    Exit Property
ErrHandler:
    Err.Raise ErrCode, TypeName(Me), "Could not set the CurrentRange to the specified Object." _
        & vbCrLf & "Ensure the NewRange argument is a Range Object and belongs to the CurrentWorksheet."
    On Error GoTo 0
End Property
