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
'@Folder("BetterArray.Tests.Dependencies.ExcelProvider")
'@ModuleDescription("An RAII container providing access to an Excel application. Dependency for 'TestModule_BetterArray'")
'@IgnoreModule FunctionReturnValueDiscarded
'@IgnoreModule ProcedureNotUsed
Option Explicit

Private Const ErrCode As Long = vbObjectError + 513

Private pExcelApplication As Object
Private pCurrentWorkbook As Object
Private pCurrentWorksheet As Object
Private pVisible As Boolean

'@Description("Constructor")
Private Sub Class_Initialize()
Attribute Class_Initialize.VB_Description = "Constructor"
    On Error GoTo ErrHandler
    Set pExcelApplication = CreateObject("Excel.Application")
    Set pCurrentWorkbook = pExcelApplication.Workbooks.Add
    Set pCurrentWorksheet = pCurrentWorkbook.Worksheets.Add
    pCurrentWorksheet.Name = TypeName(Me)
    pExcelApplication.Visible = False
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

Public Property Get Visible() As Boolean
    Visible = pVisible
End Property

Public Property Let Visible(ByVal value As Boolean)
    pVisible = value
    If Not pExcelApplication Is Nothing Then
        pExcelApplication.Visible = pVisible
    End If
End Property

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

'@Description("Returns the Excel workbook object")
Public Property Get CurrentWorksheet() As Object
Attribute CurrentWorksheet.VB_Description = "Returns the Excel workbook object"
    Set CurrentWorksheet = pCurrentWorksheet
End Property
