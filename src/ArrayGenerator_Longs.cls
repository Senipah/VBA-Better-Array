VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayGenerator_Longs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IArrayGenerator

Private i As Long
Private j As Long

Private Function rand(Max As Long, Optional Min As Long = 0) As Long
    Randomize: rand = Int(Rnd * (Max - Min) + 1) + Min
End Function

Private Function getRandomValue() As Long
    getRandomValue = rand(-1000, 1000)
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



