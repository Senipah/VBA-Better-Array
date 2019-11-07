---
title: ExtractSegment
parent: Methods
grand_parent: API
---

# ExtractSegment

## Description
The `ExtractSegment()` method extracts the specified segment of an array. If the current instance stores a two-dimensional array, you can enter a row or column index to return the specified segment of the array as a one-dimension array. If both row and column index arguments are provided the element stored at the intersection will be returned (wrapped in an array if the element is not already an array). 

If the stored array is one-dimension and both column and row arguments are provided, the element at the row index will be returned (encased in an array). If just a row or just a column index are provided the element at whichever index has been provided will be returned (encased in an array).

## Syntax

*expression*.**ExtractSegment**([*rowIndex*], [*columnIndex*])

### Parameters

Name 
: `rowIndex`

Type
: `Long`

Necessity
: Optional

Description
: The index of the row to be extracted. 

---

Name 
: `columnIndex`

Type
: `Long`

Necessity
: Optional

Description
: The index of the column to be extracted. 

### Returns

Type
: `Variant()`

Description
: A variant array containing the extracted segment.

## Example

```vb
Public Sub ExtractSegmentExample()
    Dim MultiDimensionArray(1 To 2, 1 To 2) As Variant
    Dim MyArray As BetterArray
    Dim result() As Variant
    
    MultiDimensionArray(1, 1) = "Foo"
    MultiDimensionArray(1, 2) = 1
    MultiDimensionArray(2, 1) = "Bar"
    MultiDimensionArray(2, 2) = 2
    
    Set MyArray = New BetterArray
    MyArray.Items = MultiDimensionArray
    result = MyArray.ExtractSegment(columnIndex:=2)

    ' expected output:
    ' result is a one-dimension array with the values: 1, 2
End Sub
```


[Back to Docs](https://senipah.github.io/VBA-Better-Array/)