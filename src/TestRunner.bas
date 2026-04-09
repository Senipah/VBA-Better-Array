Attribute VB_Name = "TestRunner"
Option Explicit

Private Const RUNNER_PREFIX As String = "[NoRDTests]"
Private Const RUNNER_ERROR As Long = vbObjectError + 514

Public Sub RunAllTests( _
    Optional ByVal ModuleFilter As String = vbNullString, _
    Optional ByVal TestNamePattern As String = "*" _
)
    Dim Report As String
    Report = RunAllTests_Report(ModuleFilter, TestNamePattern)
    Debug.Print Report
    If InStr(1, Report, "Failed: 0", vbTextCompare) = 0 Then
        Err.Raise RUNNER_ERROR, "TestRunner", Report
    End If
End Sub

Public Function RunAllTests_Report( _
    Optional ByVal ModuleFilter As String = vbNullString, _
    Optional ByVal TestNamePattern As String = "*" _
) As String
    Dim StartedAt As Single
    Dim TotalTests As Long
    Dim PassedTests As Long
    Dim FailedTests As Long
    Dim ModuleNames As Collection
    Dim ModuleName As Variant
    Dim HasMatch As Boolean
    Dim Failures As Collection

    If Len(TestNamePattern) = 0 Then TestNamePattern = "*"

    Set ModuleNames = GetTestModuleNames()
    Set Failures = New Collection
    StartedAt = Timer
    Debug.Print RUNNER_PREFIX & " Starting test run..."

    For Each ModuleName In ModuleNames
        If MatchesPattern(CStr(ModuleName), ModuleFilter) Then
            HasMatch = True
            RunTestsInModule CStr(ModuleName), TestNamePattern, TotalTests, PassedTests, FailedTests, Failures
        End If
    Next

    If Not HasMatch Then
        RunAllTests_Report = RUNNER_PREFIX & " No modules matched filter '" & ModuleFilter & "'."
        Exit Function
    End If

    RunAllTests_Report = BuildRunReport( _
        TotalTests, PassedTests, FailedTests, ElapsedSeconds(StartedAt), Failures _
    )
End Function

Private Sub RunTestsInModule( _
    ByVal ModuleName As String, _
    ByVal TestNamePattern As String, _
    ByRef TotalTests As Long, _
    ByRef PassedTests As Long, _
    ByRef FailedTests As Long, _
    ByRef Failures As Collection _
)
    Dim TestMethods As Collection
    Dim TestName As Variant
    Dim ErrorMessage As String
    Dim FailureDetail As String

    Set TestMethods = GetRunnableTestMethods(ModuleName)
    If TestMethods.Count = 0 Then Exit Sub

    Debug.Print RUNNER_PREFIX & " Module: " & ModuleName

    If Not TryRunProcedure(ModuleName, "ModuleInitialize", ErrorMessage) Then
        Debug.Print "  FAIL ModuleInitialize: " & ErrorMessage
        TotalTests = TotalTests + 1
        FailedTests = FailedTests + 1
        Failures.Add ModuleName & ".ModuleInitialize :: " & ErrorMessage
        Exit Sub
    End If

    For Each TestName In TestMethods
        If MatchesPattern(CStr(TestName), TestNamePattern) Then
            TotalTests = TotalTests + 1
            If RunSingleTest(ModuleName, CStr(TestName), FailureDetail) Then
                PassedTests = PassedTests + 1
                Debug.Print "  PASS " & TestName
            Else
                FailedTests = FailedTests + 1
                Debug.Print "  FAIL " & TestName & " :: " & FailureDetail
                Failures.Add ModuleName & "." & CStr(TestName) & " :: " & FailureDetail
            End If
        End If
    Next

    If Not TryRunProcedure(ModuleName, "ModuleCleanup", ErrorMessage) Then
        Debug.Print "  FAIL ModuleCleanup: " & ErrorMessage
        TotalTests = TotalTests + 1
        FailedTests = FailedTests + 1
        Failures.Add ModuleName & ".ModuleCleanup :: " & ErrorMessage
    End If
End Sub

Private Function RunSingleTest( _
    ByVal ModuleName As String, _
    ByVal TestName As String, _
    ByRef FailureDetail As String _
) As Boolean
    Dim InitMessage As String
    Dim TestMessage As String
    Dim CleanupMessage As String

    ResetTestFailureState

    If Not TryRunProcedure(ModuleName, "TestInitialize", InitMessage) Then
        FailureDetail = "TestInitialize failed: " & InitMessage
        TryRunProcedure ModuleName, "TestCleanup", CleanupMessage
        Exit Function
    End If
    If TestFailed() Then
        FailureDetail = "TestInitialize assertion failed: " & TestFailureMessage()
        TryRunProcedure ModuleName, "TestCleanup", CleanupMessage
        Exit Function
    End If

    If Not TryRunProcedure(ModuleName, TestName, TestMessage) Then
        FailureDetail = TestMessage
        TryRunProcedure ModuleName, "TestCleanup", CleanupMessage
        Exit Function
    End If
    If TestFailed() Then
        FailureDetail = TestFailureMessage()
        TryRunProcedure ModuleName, "TestCleanup", CleanupMessage
        Exit Function
    End If

    If Not TryRunProcedure(ModuleName, "TestCleanup", CleanupMessage) Then
        FailureDetail = "TestCleanup failed: " & CleanupMessage
        Exit Function
    End If
    If TestFailed() Then
        FailureDetail = "TestCleanup assertion failed: " & TestFailureMessage()
        Exit Function
    End If

    RunSingleTest = True
End Function

Private Function TryRunProcedure( _
    ByVal ModuleName As String, _
    ByVal ProcedureName As String, _
    ByRef ErrorMessage As String _
) As Boolean
    On Error GoTo RunFail

    Application.Run "'" & ThisWorkbook.Name & "'!" & ModuleName & "." & ProcedureName
    TryRunProcedure = True
    Exit Function

RunFail:
    ErrorMessage = "#" & Err.Number & " - " & Err.Description
End Function

Private Function GetTestModuleNames() As Collection
    Dim Result As Collection
    Set Result = New Collection

    If Not LoadTestModuleNamesFromVBProject(Result) Then
        LoadTestModuleNamesFromSourceFolder Result
    End If

    Set GetTestModuleNames = Result
End Function

Private Function LoadTestModuleNamesFromVBProject(ByRef Result As Collection) As Boolean
    Dim VBProject As Object
    Dim Component As Object

    On Error GoTo AccessDenied
    Set VBProject = Application.VBE.ActiveVBProject

    For Each Component In VBProject.VBComponents
        ' 1 = vbext_ct_StdModule
        If Component.Type = 1 And Component.Name Like "TestModule_*" Then
            AddUnique Result, Component.Name
        End If
    Next

    LoadTestModuleNamesFromVBProject = (Result.Count > 0)
    Exit Function

AccessDenied:
    LoadTestModuleNamesFromVBProject = False
End Function

Private Sub LoadTestModuleNamesFromSourceFolder(ByRef Result As Collection)
    Dim SourceFolder As String
    Dim FileName As String
    Dim ModuleName As String

    SourceFolder = ResolveSourceFolder()
    If Len(SourceFolder) = 0 Then Exit Sub

    FileName = Dir$(SourceFolder & Application.PathSeparator & "TestModule_*.bas")
    Do While Len(FileName) > 0
        ModuleName = Left$(FileName, Len(FileName) - 4)
        AddUnique Result, ModuleName
        FileName = Dir$
    Loop
End Sub

Private Function GetRunnableTestMethods(ByVal ModuleName As String) As Collection
    Dim Result As Collection
    Dim VBProject As Object
    Dim Component As Object
    Dim SourceFilePath As String

    Set Result = New Collection

    On Error GoTo UseSourceFile
    Set VBProject = Application.VBE.ActiveVBProject
    Set Component = VBProject.VBComponents(ModuleName)
    ParseRunnableMethodsFromCodeModule Component.CodeModule, Result
    Set GetRunnableTestMethods = Result
    Exit Function

UseSourceFile:
    SourceFilePath = ResolveSourceFolder()
    If Len(SourceFilePath) > 0 Then
        SourceFilePath = SourceFilePath & Application.PathSeparator & ModuleName & ".bas"
        If FileExists(SourceFilePath) Then ParseRunnableMethodsFromFile SourceFilePath, Result
    End If
    Set GetRunnableTestMethods = Result
End Function

Private Sub ParseRunnableMethodsFromCodeModule(ByVal CodeModule As Object, ByRef Result As Collection)
    Dim i As Long
    Dim lineText As String
    Dim MethodName As String

    For i = 1 To CodeModule.CountOfLines
        lineText = Trim$(CodeModule.Lines(i, 1))
        MethodName = ExtractSubName(lineText)
        If Len(MethodName) > 0 Then
            If Not IsLifecycleMethod(MethodName) Then
                AddUnique Result, MethodName
            End If
        End If
    Next
End Sub

Private Sub ParseRunnableMethodsFromFile(ByVal FilePath As String, ByRef Result As Collection)
    Dim FileNum As Integer
    Dim RawText As String
    Dim lines As Variant
    Dim i As Long
    Dim lineText As String
    Dim MethodName As String

    On Error GoTo FileReadFail
    FileNum = FreeFile
    Open FilePath For Input As #FileNum
    RawText = Input$(LOF(FileNum), FileNum)
    Close #FileNum

    RawText = Replace(RawText, vbCrLf, vbLf)
    lines = Split(RawText, vbLf)

    For i = LBound(lines) To UBound(lines)
        lineText = Trim$(CStr(lines(i)))
        MethodName = ExtractSubName(lineText)
        If Len(MethodName) > 0 Then
            If Not IsLifecycleMethod(MethodName) Then
                AddUnique Result, MethodName
            End If
        End If
    Next
    Exit Sub

FileReadFail:
    On Error Resume Next
    If FileNum > 0 Then Close #FileNum
End Sub

Private Function IsLifecycleMethod(ByVal MethodName As String) As Boolean
    Select Case LCase$(MethodName)
    Case "moduleinitialize", "modulecleanup", "testinitialize", "testcleanup"
        IsLifecycleMethod = True
    End Select
End Function

Private Function ExtractSubName(ByVal SignatureLine As String) As String
    Dim Signature As String
    Dim SignatureLower As String
    Dim PrefixLength As Long
    Dim OpenParenIndex As Long
    Dim Candidate As String

    Signature = Trim$(SignatureLine)
    SignatureLower = LCase$(Signature)

    If Left$(SignatureLower, 11) = "public sub " Then
        PrefixLength = 11
    ElseIf Left$(SignatureLower, 12) = "private sub " Then
        PrefixLength = 12
    Else
        Exit Function
    End If

    OpenParenIndex = InStr(1, Signature, "(")
    If OpenParenIndex <= PrefixLength Then Exit Function

    Candidate = Trim$(Mid$(Signature, PrefixLength + 1, OpenParenIndex - PrefixLength - 1))
    If Len(Candidate) = 0 Then Exit Function
    ExtractSubName = Candidate
End Function

Private Sub AddUnique(ByRef Values As Collection, ByVal ItemValue As String)
    On Error GoTo AlreadyExists
    Values.Add ItemValue, LCase$(ItemValue)
    Exit Sub

AlreadyExists:
End Sub

Private Function ResolveSourceFolder() As String
    Dim Candidate As String
    Dim ParentFolder As String

    Candidate = ThisWorkbook.Path & Application.PathSeparator & "src"
    If DirectoryExists(Candidate) Then
        ResolveSourceFolder = Candidate
        Exit Function
    End If

    ParentFolder = ParentPath(ThisWorkbook.Path)
    If Len(ParentFolder) > 0 Then
        Candidate = ParentFolder & Application.PathSeparator & "src"
        If DirectoryExists(Candidate) Then
            ResolveSourceFolder = Candidate
            Exit Function
        End If
    End If
End Function

Private Function ParentPath(ByVal SourcePath As String) As String
    Dim LastSeparatorIndex As Long
    LastSeparatorIndex = InStrRev(SourcePath, Application.PathSeparator)
    If LastSeparatorIndex > 0 Then
        ParentPath = Left$(SourcePath, LastSeparatorIndex - 1)
    End If
End Function

Private Function DirectoryExists(ByVal FolderPath As String) As Boolean
    On Error GoTo Missing
    DirectoryExists = ((GetAttr(FolderPath) And vbDirectory) = vbDirectory)
    Exit Function
Missing:
    DirectoryExists = False
End Function

Private Function FileExists(ByVal FilePath As String) As Boolean
    On Error GoTo Missing
    FileExists = (Len(Dir$(FilePath)) > 0)
    Exit Function
Missing:
    FileExists = False
End Function

Private Function MatchesPattern(ByVal Value As String, ByVal Pattern As String) As Boolean
    If Len(Pattern) = 0 Then
        MatchesPattern = True
    Else
        MatchesPattern = (LCase$(Value) Like LCase$(Pattern))
    End If
End Function

Private Function BuildRunReport( _
    ByVal TotalTests As Long, _
    ByVal PassedTests As Long, _
    ByVal FailedTests As Long, _
    ByVal DurationSeconds As Double, _
    ByRef Failures As Collection _
) As String
    Dim Report As String
    Report = RUNNER_PREFIX & " Done." & vbCrLf & _
        RUNNER_PREFIX & " Total: " & TotalTests & ", Passed: " & PassedTests & ", Failed: " & FailedTests & vbCrLf & _
        RUNNER_PREFIX & " Duration: " & Format$(DurationSeconds, "0.00") & "s"

    If FailedTests > 0 Then
        Report = Report & vbCrLf & RUNNER_PREFIX & " Failures:" & vbCrLf & FailureItemsToString(Failures)
    End If

    BuildRunReport = Report
End Function

Private Function FailureItemsToString(ByRef Failures As Collection) As String
    Dim i As Long
    Dim Buffer As String
    Dim ItemText As String

    For i = 1 To Failures.Count
        ItemText = CStr(Failures.Item(i))
        Buffer = Buffer & i & ". " & ItemText & vbCrLf
    Next
    If Len(Buffer) > 0 Then
        Buffer = Left$(Buffer, Len(Buffer) - Len(vbCrLf))
    End If
    FailureItemsToString = Buffer
End Function

Private Function ElapsedSeconds(ByVal StartAt As Single) As Double
    If Timer >= StartAt Then
        ElapsedSeconds = Timer - StartAt
    Else
        ElapsedSeconds = (86400# - StartAt) + Timer
    End If
End Function
