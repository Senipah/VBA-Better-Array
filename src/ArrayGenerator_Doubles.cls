VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayGenerator_Doubles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IArrayGenerator

'@Folder("Tests.Dependencies.IArrayGenerator.Implementations")

Private i As Long
Private j As Long

Private Function rand(ByVal Max As Long, Optional ByVal Min As Long = 0) As Long
    Randomize: rand = Int(Rnd * (Max - Min) + 1) + Min
End Function

Private Function getRandomValue() As Double
    getRandomValue = rand(-1000, 1000) * Rnd
End Function

Public Function IArrayGenerator_getOneDimensionalArray(Length As Long) As Variant()
    Dim arr() As Variant
    ReDim arr(Length)
    For i = 0 To Length
        arr(i) = getRandomValue()
    Next
    IArrayGenerator_getOneDimensionalArray = arr
End Function

Public Function IArrayGenerator_getMultidimensionalArray(Length As Long) As Variant()
    Dim arr() As Variant
    ReDim arr(Length, Length)
    For i = 0 To Length
        For j = 0 To Length
            arr(i, j) = getRandomValue()
        Next
    Next
    IArrayGenerator_getMultidimensionalArray = arr
End Function

Public Function IArrayGenerator_getJaggedArray(Length As Long) As Variant()
    Dim parentArray() As Variant
    Dim childArray() As Variant
    ReDim parentArray(Length)
    ReDim childArray(Length)
    For i = 0 To Length
        For j = 0 To Length
            childArray(j) = getRandomValue()
        Next
        parentArray(i) = childArray
    Next
    IArrayGenerator_getJaggedArray = parentArray
End Function


