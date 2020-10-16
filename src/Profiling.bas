Attribute VB_Name = "Profiling"
Option Explicit
Option Private Module

'@Folder("BetterArray.Tests.Performance")

'@IgnoreModule FunctionReturnValueNotUsed
'@IgnoreModule FunctionReturnValueDiscarded
Private Function formatDescriptor(ByVal descriptor As String) As String
    Dim result As String
    Dim corner As String
    Dim vertice As String
    Dim horizon As String
    corner = "+"
    vertice = "|"
    horizon = "-"
    result = corner & String(Len(descriptor) + 2, horizon) & corner & vbCrLf _
           & vertice & " " & descriptor & " " & vertice & vbCrLf _
           & corner & String(Len(descriptor) + 2, horizon) & corner
    formatDescriptor = result
End Function

Public Sub PushingScalar()
    Const maxEntries As Long = 1000000
    Const descriptor As String = "Pushing {count} Scalar Values."
    
    Dim i As Long
    Dim betterArrayTime As Double
    Dim manualTime As Double
    
    i = 10
    Do While i <= maxEntries
        Debug.Print formatDescriptor(Replace(descriptor, "{count}", CStr(i)))
        manualTime = PushingScalarByRedim(i)
        betterArrayTime = PushingScalarByBetterArray(i)
        RatePerformance manualTime, betterArrayTime
        i = i * 10
        DoEvents
    Loop
    
End Sub

Private Sub RatePerformance(ByVal manualTime As Double, ByVal betterArrayTime As Double)
    Const descriptor As String = "Time taken with "
    Const resultStart As String = "BetterArray is "
    Const resultEnd As String = " Than the manual method."
    Dim diff As Double
    
    
    diff = manualTime - betterArrayTime
    If diff <> 0 And betterArrayTime <> 0 Then diff = diff / betterArrayTime
    Debug.Print descriptor & "manual method: " & manualTime
    Debug.Print descriptor & "BetterArray: " & betterArrayTime
    If diff <> 0 Then
        Debug.Print resultStart _
                    & Format$(Abs(diff), "Percent") _
                    & IIf(diff > 0, " faster", " slower") _
                    & resultEnd
    Else
        Debug.Print "Effectively same speed."
    End If
End Sub

Private Function PushingScalarByRedim(ByVal count As Long) As Double
    Dim SUT() As Variant
    Dim startTime As Double
    Dim i As Long
    startTime = Timer
    
    For i = 0 To count - 1
        ReDim Preserve SUT(i)
        SUT(i) = i
    Next
    
    PushingScalarByRedim = Timer - startTime
End Function

Private Function PushingScalarByBetterArray(ByVal count As Long) As Double
    Dim SUT As BetterArray
    Dim startTime As Double
    Dim i As Long
    startTime = Timer
    Set SUT = New BetterArray
    For i = 0 To count - 1
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
        Debug.Print formatDescriptor(Replace(descriptor, "{count}", CStr(i)))
        manualTime = PushingArraysByRedim(i)
        betterArrayTime = PushingArraysByBetterArray(i)
        RatePerformance manualTime, betterArrayTime
        i = i * 10
        DoEvents
    Loop
End Sub

Private Function PushingArraysByRedim(ByVal count As Long) As Double
    Dim SUT() As Variant
    Dim startTime As Double
    Dim i As Long
    startTime = Timer
    
    For i = 0 To count - 1
        ReDim Preserve SUT(i)
        SUT(i) = Array(1, 2, 3)
    Next
    
    PushingArraysByRedim = Timer - startTime
End Function

Private Function PushingArraysByBetterArray(ByVal count As Long) As Double
    Dim SUT As BetterArray
    Dim startTime As Double
    '@Ignore VariableNotUsed
    Dim i As Long
    startTime = Timer
    Set SUT = New BetterArray
    For i = 0 To count - 1
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
        Debug.Print formatDescriptor(Replace(descriptor, "{count}", CStr(i)))
        manualTime = TransposingByTranspose(i)
        betterArrayTime = TransposingByBetterArray(i)
        RatePerformance manualTime, betterArrayTime
        i = i * 10
        DoEvents
    Loop

End Sub

Private Function TransposingByTranspose(ByVal count As Long) As Double
    Dim exl As ExcelProvider
    Dim destination As Range
    Dim SUT() As Variant
    Dim startTime As Double
    Dim i As Long
    
    For i = 0 To count - 1
        ReDim Preserve SUT(i)
        SUT(i) = Array(1, 2, 3)
    Next
    Set exl = New ExcelProvider
    Set destination = exl.CurrentWorksheet.Range("A1")
    startTime = Timer
    '@Ignore ImplicitDefaultMemberAccess
    destination.Resize(count, 3) = WorksheetFunction.Transpose(WorksheetFunction.Transpose(SUT))
    exl.Visible = True
    TransposingByTranspose = Timer - startTime
End Function


Private Function TransposingByBetterArray(ByVal count As Long) As Double
    Dim exl As ExcelProvider
    Dim destination As Object
    Dim SUT As BetterArray
    Dim startTime As Double
    '@Ignore VariableNotUsed
    Dim i As Long
    Set SUT = New BetterArray
    For i = 0 To count - 1
        SUT.Push Array(1, 2, 3)
    Next
    Set exl = New ExcelProvider
    Set destination = exl.CurrentWorksheet.Range("A1")
    startTime = Timer
    SUT.ToExcelRange destination
    exl.Visible = True
    TransposingByBetterArray = Timer - startTime
End Function

Public Sub CSV_Test()
    Const DATA_DIR As String = "csv_data"
    Const SLUG As String = " Sales Records.csv"
    Dim i As Long
    Dim startTime As Single
    Dim endTime As Single
    Dim basePath As String
    Dim filePath As String
    Dim fileName As String
    Dim recordCounts() As Variant
    Dim SUT As BetterArray
    Set SUT = New BetterArray
    
    basePath = ThisWorkbook.path
    
    recordCounts = Array(10000)
            
    For i = LBound(recordCounts) To UBound(recordCounts)
        fileName = recordCounts(i) & SLUG
        filePath = JoinPath(basePath, DATA_DIR, fileName)
        Debug.Print formatDescriptor("Reading: " & fileName)
        startTime = Timer
        SUT.FromCSVFile filePath
        endTime = Timer
        Debug.Print "Time taken: " & endTime - startTime
        SUT.ToExcelRange ThisWorkbook.Sheets.Add.Range("A1")
    Next
    
End Sub

Public Sub CSV_Profiling()
    Const DATA_DIR As String = "csv_data"
    Const SLUG As String = " Sales Records.csv"
    Dim i As Long
    Dim startTime As Single
    Dim endTime As Single
    Dim basePath As String
    Dim filePath As String
    Dim fileName As String
    Dim recordCounts() As Variant
    Dim SUT As BetterArray
    Set SUT = New BetterArray
    
    basePath = ThisWorkbook.path
    
    recordCounts = Array(10000, 100000, 1500000)
            
    For i = LBound(recordCounts) To UBound(recordCounts)
        fileName = recordCounts(i) & SLUG
        filePath = JoinPath(basePath, DATA_DIR, fileName)
        Debug.Print formatDescriptor("Reading: " & fileName)
        startTime = Timer
        SUT.FromCSVFile filePath
        endTime = Timer
        Debug.Print "Time taken: " & endTime - startTime
    Next
    
End Sub

Private Function JoinPath(ParamArray args() As Variant) As String
    JoinPath = Strings.Join(CVar(args), "\")
End Function

