Attribute VB_Name = "Profiling"
Option Explicit
Option Private Module

'@Folder("Tests.Performance")

Public Sub PushingScalar()
    Const maxEntries As Long = 1000000
    Const descriptor As String = "Pushing {count} Scalar Values."
    
    Dim i As Long
    Dim betterArrayTime As Double
    Dim manualTime As Double
    
    i = 10
    Do While i <= maxEntries
        Debug.Print Replace(descriptor, "{count}", CStr(i))
        manualTime = PushingScalarByRedim(i)
        betterArrayTime = PushingScalarByBetterArray(i)
        RatePerformance manualTime, betterArrayTime
        i = i * 10
        DoEvents
    Loop
    
End Sub

Private Sub RatePerformance(ByVal manualTime As Double, ByVal betterArrayTime As Double)
    Const descriptor As String = "Time taken with "
    Const resultStart As String = "Better Array is "
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
    Dim sut() As Variant
    Dim startTime As Double
    Dim i As Long
    startTime = Timer
    
    For i = 0 To count - 1
        ReDim Preserve sut(i)
        sut(i) = i
    Next
    
    PushingScalarByRedim = Timer - startTime
End Function

Private Function PushingScalarByBetterArray(ByVal count As Long) As Double
    Dim sut As BetterArray
    Dim startTime As Double
    Dim i As Long
    startTime = Timer
    Set sut = New BetterArray
    For i = 0 To count - 1
        sut.Push i
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
        Debug.Print Replace(descriptor, "{count}", CStr(i))
        manualTime = PushingArraysByRedim(i)
        betterArrayTime = PushingArraysByBetterArray(i)
        RatePerformance manualTime, betterArrayTime
        i = i * 10
        DoEvents
    Loop
End Sub

Private Function PushingArraysByRedim(ByVal count As Long) As Double
    Dim sut() As Variant
    Dim startTime As Double
    Dim i As Long
    startTime = Timer
    
    For i = 0 To count - 1
        ReDim Preserve sut(i)
        sut(i) = Array(1, 2, 3)
    Next
    
    PushingArraysByRedim = Timer - startTime
End Function

Private Function PushingArraysByBetterArray(ByVal count As Long) As Double
    Dim sut As BetterArray
    Dim startTime As Double
    '@Ignore VariableNotUsed
    Dim i As Long
    startTime = Timer
    Set sut = New BetterArray
    For i = 0 To count - 1
        sut.Push Array(1, 2, 3)
    Next
    
    PushingArraysByBetterArray = Timer - startTime
End Function


Public Sub TransposingJaggedToExcel()
    Const maxEntries As Long = 100000
    Const descriptor As String = "Transpsing To Excel {count} Arrays."
    
    Dim i As Long
    Dim betterArrayTime As Double
    Dim manualTime As Double
    
    i = 10
    Do While i <= maxEntries
        Debug.Print Replace(descriptor, "{count}", CStr(i))
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
    Dim sut() As Variant
    Dim startTime As Double
    Dim i As Long
    
    For i = 0 To count - 1
        ReDim Preserve sut(i)
        sut(i) = Array(1, 2, 3)
    Next
    Set exl = New ExcelProvider
    Set destination = exl.CurrentWorksheet.Range("A1")
    startTime = Timer
    '@Ignore ImplicitDefaultMemberAccess
    destination.Resize(count, 3) = WorksheetFunction.Transpose(WorksheetFunction.Transpose(sut))
    exl.Visible = True
    TransposingByTranspose = Timer - startTime
End Function


Private Function TransposingByBetterArray(ByVal count As Long) As Double
    Dim exl As ExcelProvider
    Dim destination As Object
    Dim sut As BetterArray
    Dim startTime As Double
    '@Ignore VariableNotUsed
    Dim i As Long
    Set sut = New BetterArray
    For i = 0 To count - 1
        sut.Push Array(1, 2, 3)
    Next
    Set exl = New ExcelProvider
    Set destination = exl.CurrentWorksheet.Range("A1")
    startTime = Timer
    sut.ToExcelRange destination
    exl.Visible = True
    TransposingByBetterArray = Timer - startTime
End Function






