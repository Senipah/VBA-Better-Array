Attribute VB_Name = "Profiling"
'@Folder("VBABetterArray.Tests.Misc")

Option Explicit
Option Private Module

'@IgnoreModule FunctionReturnValueNotUsed
'@IgnoreModule FunctionReturnValueDiscarded
'@IgnoreModule ProcedureNotUsed

Public Sub PushingScalar()
    Const maxEntries As Long = 1000000
    Const descriptor As String = "Pushing {count} Scalar Values."
    
    Dim i As Long
    Dim betterArrayTime As Double
    Dim manualTime As Double
    
    i = 10
    Do While i <= maxEntries
        Debug.Print ConsoleHeader(Replace(descriptor, "{count}", CStr(i)))
        manualTime = PushingScalarByRedim(i)
        betterArrayTime = PushingScalarByBetterArray(i)
        RatePerformance manualTime, betterArrayTime
        i = i * 10
        DoEvents
    Loop
    
End Sub

Private Function PushingScalarByRedim(ByVal Count As Long) As Double
    Dim SUT() As Variant
    Dim startTime As Double
    Dim i As Long
    startTime = Timer
    
    For i = 0 To Count - 1
        ReDim Preserve SUT(i)
        SUT(i) = i
    Next
    
    PushingScalarByRedim = Timer - startTime
End Function

Private Function PushingScalarByBetterArray(ByVal Count As Long) As Double
    Dim SUT As BetterArray
    Dim startTime As Double
    Dim i As Long
    startTime = Timer
    Set SUT = New BetterArray
    For i = 0 To Count - 1
        SUT.Push i
    Next
    
    PushingScalarByBetterArray = Timer - startTime
End Function

Public Sub PushingArrays()
    Const maxEntries As Long = 1000000
    Const descriptor As String = "Pushing {count} Arrays."
    
    Dim i As Long
    Dim betterArrayTime As Double
    Dim manualTime As Double
    
    i = 10
    Do While i <= maxEntries
        Debug.Print ConsoleHeader(Replace(descriptor, "{count}", CStr(i)))
        manualTime = PushingArraysByRedim(i)
        betterArrayTime = PushingArraysByBetterArray(i)
        RatePerformance manualTime, betterArrayTime
        i = i * 10
        DoEvents
    Loop
End Sub

Private Function PushingArraysByRedim(ByVal Count As Long) As Double
    Dim SUT() As Variant
    Dim startTime As Double
    Dim i As Long
    startTime = Timer
    
    For i = 0 To Count - 1
        ReDim Preserve SUT(i)
        SUT(i) = Array(1, 2, 3)
    Next
    
    PushingArraysByRedim = Timer - startTime
End Function

Private Function PushingArraysByBetterArray(ByVal Count As Long) As Double
    Dim SUT As BetterArray
    Dim startTime As Double
    '@Ignore VariableNotUsed
    Dim i As Long
    startTime = Timer
    Set SUT = New BetterArray
    For i = 0 To Count - 1
        SUT.Push Array(1, 2, 3)
    Next
    
    PushingArraysByBetterArray = Timer - startTime
End Function

Public Sub TransposingJaggedToExcel()
    Const maxEntries As Long = 100000
    Const descriptor As String = "Transposing To Excel {count} Arrays."
    
    Dim i As Long
    Dim betterArrayTime As Double
    Dim manualTime As Double
    
    i = 10
    Do While i <= maxEntries
        Debug.Print ConsoleHeader(Replace(descriptor, "{count}", CStr(i)))
        manualTime = TransposingByTranspose(i)
        betterArrayTime = TransposingByBetterArray(i)
        RatePerformance manualTime, betterArrayTime
        i = i * 10
        DoEvents
    Loop

End Sub

Private Function TransposingByTranspose(ByVal Count As Long) As Double
    Dim exl As ExcelProvider
    Dim Destination As Range
    Dim SUT() As Variant
    Dim startTime As Double
    Dim i As Long
    
    For i = 0 To Count - 1
        ReDim Preserve SUT(i)
        SUT(i) = Array(1, 2, 3)
    Next
    Set exl = New ExcelProvider
    Set Destination = exl.CurrentWorksheet.Range("A1")
    startTime = Timer
    '@Ignore ImplicitDefaultMemberAccess
    Destination.Resize(Count, 3) = WorksheetFunction.Transpose(WorksheetFunction.Transpose(SUT))
    exl.Visible = True
    TransposingByTranspose = Timer - startTime
End Function


Private Function TransposingByBetterArray(ByVal Count As Long) As Double
    Dim exl As ExcelProvider
    Dim Destination As Object
    Dim SUT As BetterArray
    Dim startTime As Double
    '@Ignore VariableNotUsed
    Dim i As Long
    Set SUT = New BetterArray
    For i = 0 To Count - 1
        SUT.Push Array(1, 2, 3)
    Next
    Set exl = New ExcelProvider
    Set Destination = exl.CurrentWorksheet.Range("A1")
    startTime = Timer
    SUT.ToExcelRange Destination
    exl.Visible = True
    TransposingByBetterArray = Timer - startTime
End Function

Public Sub CSV_Profiling()
    Const DATA_DIR As String = "csv_data"
    Const SLUG As String = " Sales Records.csv"
    Dim i As Long
    Dim basepath As String
    Dim fileName As String
    Dim recordCounts() As Variant
    Dim SUT As BetterArray
    Set SUT = New BetterArray
    
    basepath = ThisWorkbook.Path
    
    ' recordCounts = Array(10000, 100000, 1500000)
    ' recordCounts = Array(10000, 100000)
    recordCounts = Array(1500000)
            
    For i = LBound(recordCounts) To UBound(recordCounts)
        fileName = recordCounts(i) & SLUG
        ReadCSV SUT, JoinPath(basepath, DATA_DIR), fileName
    Next
End Sub

