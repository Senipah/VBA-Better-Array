Attribute VB_Name = "SortingTests"
Option Explicit
'@Folder("BetterArray.WIP.Sorting")
'@IgnoreModule FunctionReturnValueDiscarded

Private Enum DataSizes
    Small = 10000
    Medium = 100000
    Large = 1500000
End Enum

Private Function Sum1DArray(ByRef source() As Variant) As Long
    Dim i As Long
    Dim total As Long
    For i = LBound(source) To UBound(source)
        total = total + source(i)
    Next
    Sum1DArray = total
End Function

Private Function GetCSVTestData(ByVal Size As DataSizes) As BetterArray
    Const DATA_DIR As String = "csv_data"
    Const SLUG As String = " Sales Records.csv"
    Dim filename As String
    Dim filepath As String
    Dim basePath As String
    Dim datum As BetterArray
    
    Set datum = New BetterArray
    basePath = ThisWorkbook.Path
    filename = CStr(Size) & SLUG
    filepath = JoinPath(basePath, DATA_DIR, filename)
    '@Ignore FunctionReturnValueDiscarded
    datum.FromCSVFile filepath
    Debug.Print "Parsed "; datum.Length; " rows from CSV file"
    Set GetCSVTestData = datum
End Function

Private Function ArrayFactory(ByVal ValueType As ValueTypes, ByVal ArrayType As AG_ArrayTypes, ByVal Length As Long, Optional ByVal Depth As Long) As BetterArray
    Dim Result As BetterArray
    Set Result = New BetterArray
    With New ArrayGenerator
        Result.Items = .GetArray(ValueType, ArrayType, Length, Depth)
    End With
    Set ArrayFactory = Result
End Function

Public Sub TestSmall()
    SpeedTestQSRecursive1DSML
    SpeedTestQSIterative1DSML
    SpeedTestQSRecursive2DSML
    SpeedTestQSIterative2DSML
End Sub

Public Sub TestMedium()
    SpeedTestQSRecursive1DMED
    SpeedTestQSIterative1DMED
    SpeedTestQSRecursive2DMED
    SpeedTestQSIterative2DMED
End Sub

Public Sub TestLarge()
    SpeedTestQSRecursive1DLRG
    SpeedTestQSIterative1DLRG
    SpeedTestQSRecursive2DLRG
    SpeedTestQSIterative2DLRG
End Sub

Public Sub SpeedTestQSRecursive1DSML()
    Recursive1DTest Small
End Sub

Public Sub SpeedTestQSRecursive1DMED()
    Recursive1DTest Medium
End Sub

Public Sub SpeedTestQSRecursive1DLRG()
    Recursive1DTest Large
End Sub

Public Sub SpeedTestQSIterative1DSML()
    Iterative1DTest Small
End Sub

Public Sub SpeedTestQSIterative1DMED()
    Iterative1DTest Medium
End Sub

Public Sub SpeedTestQSIterative1DLRG()
    Iterative1DTest Large
End Sub

Public Sub SpeedTestQSRecursive2DSML()
    Recursive2DTest Small
End Sub

Public Sub SpeedTestQSRecursive2DMED()
    Recursive2DTest Medium
End Sub

Public Sub SpeedTestQSRecursive2DLRG()
    Recursive2DTest Large
End Sub

Public Sub SpeedTestQSIterative2DSML()
    Iterative2DTest Small
End Sub

Public Sub SpeedTestQSIterative2DMED()
    Iterative2DTest Medium
End Sub

Public Sub SpeedTestQSIterative2DLRG()
    Iterative2DTest Large
End Sub

Private Sub Recursive1DTest(ByVal Size As DataSizes)
    Dim SUT As BetterArray
    Set SUT = ArrayFactory(AG_LONG, AG_ONEDIMENSION, Size)
    Debug.Print ConsoleHeader("Sorting 1D array of " & Size & " Rows using Recursive Quicksort")
    TestSortMethod SUT, SM_QUICKSORT_RECURSIVE
End Sub

Private Sub Iterative1DTest(ByVal Size As DataSizes)
    Dim SUT As BetterArray
    Set SUT = ArrayFactory(AG_LONG, AG_ONEDIMENSION, Size)
    Debug.Print ConsoleHeader("Sorting 1D array of " & Size & " Rows using Iterative Quicksort")
    TestSortMethod SUT, SM_QUICKSORT_ITERATIVE
End Sub

Private Sub Recursive2DTest(ByVal Size As DataSizes)
    Dim SUT As BetterArray
    Set SUT = ArrayFactory(AG_LONG, AG_JAGGED, Size)
    Debug.Print ConsoleHeader("Sorting 2D array of " & Size & " Rows using Recursive Quicksort")
    TestSortMethod SUT, SM_QUICKSORT_RECURSIVE, 1
End Sub

Private Sub Iterative2DTest(ByVal Size As DataSizes)
    Dim SUT As BetterArray
    Set SUT = ArrayFactory(AG_LONG, AG_JAGGED, Size)
    Debug.Print ConsoleHeader("Sorting 2D array of " & Size & " Rows using Iterative Quicksort")
    TestSortMethod SUT, SM_QUICKSORT_ITERATIVE, 1
End Sub

Private Sub TestSortMethod(ByVal SUT As BetterArray, ByVal Algorithm As SortMethods, Optional ByVal SortColumn As Variant)
    Dim startTime As Double
    Dim endTime As Double
    
    SUT.SortMethod = Algorithm
    startTime = Timer
    If IsMissing(SortColumn) Then
        SUT.Sort
    Else
        SUT.Sort CLng(SortColumn)
    End If
    endTime = Timer
    Debug.Print "Time Taken: "; endTime - startTime
End Sub

Public Sub CSVSortSML()
    QSRecursiveCSVSort Small
    QSIterativeCSVSort Small
End Sub

Public Sub CSVSortMED()
    QSRecursiveCSVSort Medium
    QSIterativeCSVSort Medium
End Sub

Public Sub CSVSortLRG()
    QSRecursiveCSVSort Large
    QSIterativeCSVSort Large
End Sub

Private Sub QSRecursiveCSVSort(ByVal Size As DataSizes)
    Dim SUT As BetterArray
    Debug.Print ConsoleHeader("Sorting 2D CSV array of " & Size & " Rows using Recursive Quicksort")
    Set SUT = GetCSVTestData(Size)
    TestSortMethod SUT, SM_QUICKSORT_RECURSIVE, 2
End Sub

Private Sub QSIterativeCSVSort(ByVal Size As DataSizes)
    Dim SUT As BetterArray
    Debug.Print ConsoleHeader("Sorting 2D CSV array of " & Size & " Rows using Iterative Quicksort")
    Set SUT = GetCSVTestData(Size)
    TestSortMethod SUT, SM_QUICKSORT_ITERATIVE, 2
End Sub


Public Sub SortingDriver()
    Dim Gen As ArrayGenerator
    Dim Arr As BetterArray
    Dim intArray() As Variant
    
    Dim startSum As Long
    Dim sortedSum As Long
    
    Set Gen = New ArrayGenerator
    Set Arr = New BetterArray
    
    intArray = Gen.GetArray(AG_LONG, BA_ONEDIMENSION, 10000)
    startSum = Sum1DArray(intArray)
    
    Stop
    
    TimSort intArray
    sortedSum = Sum1DArray(intArray)
    Debug.Print startSum - sortedSum
    
    'Stop
    
    Arr.Items = intArray
    Debug.Print Arr.IsSorted
    
    'Stop
    
End Sub

