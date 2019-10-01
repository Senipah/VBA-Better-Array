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

'@Folder("Tests.Dependencies.IArrayGenerator.Implementations")

Private Const VALUE_COUNT As Long = 100
Private values(VALUE_COUNT) As Variant
Private i As Long
Private j As Long

Private Sub Class_Initialize()
    
    Dim pBools As IArrayGenerator
    Set pBools = New ArrayGenerator_Booleans
    Dim BoolArray As Variant
    BoolArray = pBools.getOneDimensionalArray(VALUE_COUNT)
        
    Dim pBytes As IArrayGenerator
    Set pBytes = New ArrayGenerator_Bytes
    Dim ByteArray As Variant
    ByteArray = pBytes.getOneDimensionalArray(VALUE_COUNT)
    
    Dim pDoubles As IArrayGenerator
    Set pDoubles = New ArrayGenerator_Doubles
    Dim DoubleArray As Variant
    DoubleArray = pDoubles.getOneDimensionalArray(VALUE_COUNT)
    
    Dim pLongs As IArrayGenerator
    Set pLongs = New ArrayGenerator_Longs
    Dim LongArray As Variant
    LongArray = pLongs.getOneDimensionalArray(VALUE_COUNT)
    
    Dim pStrings As IArrayGenerator
    Set pStrings = New ArrayGenerator_Strings
    Dim StringArray As Variant
    StringArray = pStrings.getOneDimensionalArray(VALUE_COUNT)
    
    Dim randType As Long
    
    For i = 0 To VALUE_COUNT
        randType = rand(4)
        Select Case randType
            Case 0
                values(i) = BoolArray(rand(VALUE_COUNT))
            Case 1
                values(i) = ByteArray(rand(VALUE_COUNT))
            Case 2
                values(i) = DoubleArray(rand(VALUE_COUNT))
            Case 3
                values(i) = LongArray(rand(VALUE_COUNT))
            Case 4
                values(i) = StringArray(rand(VALUE_COUNT))
        End Select
    Next
    
End Sub

Private Function rand(ByVal Max As Long, Optional ByVal Min As Long = 0) As Long
    Randomize: rand = Int(Rnd * (Max - Min) + 1) + Min
End Function

Private Function getRandomValue() As Variant
    getRandomValue = values(rand(LBound(values), UBound(values)))
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

