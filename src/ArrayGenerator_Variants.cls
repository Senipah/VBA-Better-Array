VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayGenerator_Variants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IArrayGenerator

Private Const nVals As Long = 100
Private values(nVals) As Variant
Private i As Long
Private j As Long

Private Sub Class_Initialize()
    
    Dim pBools As IArrayGenerator
    Set pBools = New ArrayGenerator_Booleans
    Dim BoolArray As Variant
    BoolArray = pBools.getOneDimensionalArray(nVals)
        
    Dim pBytes As IArrayGenerator
    Set pBytes = New ArrayGenerator_Bytes
    Dim ByteArray As Variant
    ByteArray = pBytes.getOneDimensionalArray(nVals)
    
    Dim pDoubles As IArrayGenerator
    Set pDoubles = New ArrayGenerator_Doubles
    Dim DoubleArray As Variant
    DoubleArray = pDoubles.getOneDimensionalArray(nVals)
    
    Dim pLongs As IArrayGenerator
    Set pLongs = New ArrayGenerator_Longs
    Dim LongArray As Variant
    LongArray = pLongs.getOneDimensionalArray(nVals)
    
    Dim pStrings As IArrayGenerator
    Set pStrings = New ArrayGenerator_Strings
    Dim StringArray As Variant
    StringArray = pStrings.getOneDimensionalArray(nVals)
    
    Dim randType As Long
    
    For i = 0 To nVals
        randType = rand(4)
        Select Case randType
            Case 0
                values(i) = BoolArray(rand(nVals))
            Case 1
                values(i) = ByteArray(rand(nVals))
            Case 2
                values(i) = DoubleArray(rand(nVals))
            Case 3
                values(i) = LongArray(rand(nVals))
            Case 4
                values(i) = StringArray(rand(nVals))
        End Select
    Next
    
End Sub

Private Function rand(Max As Long, Optional Min As Long = 0) As Long
    Randomize: rand = Int(Rnd * (Max - Min) + 1) + Min
End Function

Private Function getRandomValue() As Variant
    getRandomValue = values(rand(LBound(values), UBound(values)))
End Function

Public Function IArrayGenerator_getOneDimensionalArray(length As Long) As Variant()
    Dim arr() As Variant
    ReDim arr(length)
    For i = 0 To length
        arr(i) = getRandomValue()
    Next
    IArrayGenerator_getOneDimensionalArray = arr
End Function

Public Function IArrayGenerator_getMultidimensionalArray(length As Long) As Variant()
    Dim arr() As Variant
    ReDim arr(length, length)
    For i = 0 To length
        For j = 0 To length
            arr(i, j) = getRandomValue()
        Next
    Next
    IArrayGenerator_getMultidimensionalArray = arr
End Function

Public Function IArrayGenerator_getJaggedArray(length As Long) As Variant()
    Dim arr1() As Variant
    Dim arr2() As Variant
    ReDim arr1(length)
    ReDim arr2(length)
    For i = 0 To length
        For j = 0 To length
            arr2(j) = getRandomValue()
        Next
        arr1(i) = arr2
    Next
    IArrayGenerator_getJaggedArray = arr1
End Function


