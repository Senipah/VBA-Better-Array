---
title: Unique
parent: Methods
grand_parent: API
---

# Unique

## Description
The `Unique()` method removes all duplicate elements from the outermost array such that only unique elements remain.

## Syntax

*expression*.**Unique**([*ColumnIndex*])

### Parameters

Name
: `ColumnIndex`

Type
: `Long`

Necessity
: Optional

Description
: A base-1 index of the column in a jagged or multi-dimension array with 2 dimensions to filter by unique values. If no column index is provided and the array is jagged then Unique will compare all elements in nested arrays for equality when determining which nested arrays are Unique. If the `ColumnIndex` is greater than the max length of the arrays at the second dimension the first column at that dimension will be used.

### Returns

Type
: `BetterArray` / `Object`

Description
: The current instance of the BetterArray object with any duplicate elements removed.

## Example

```vb
Public Sub UniqueExample()
    Dim result() As Variant
    Dim MyArray As BetterArray
    Set MyArray = New BetterArray

    MyArray.Push "Foo", "Foo", "Bar", "Foo", "Buzz", "Bar"
    MyArray.Unique
    result = MyArray.Items
    ' expected output:
    ' result is an array containing "Foo","Bar","Buzz"
End Sub
```

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)
