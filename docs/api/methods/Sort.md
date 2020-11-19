---
title: Sort
parent: Methods
grand_parent: API
---

# Sort

## Description

The `Sort()` method sorts the stored array in ascending order.

The sorting algorithm used depends on the value of the [SortMethod](https://senipah.github.io/VBA-Better-Array/api/properties/sort_method/SortMethod.html) property, which must be one of the [SortMethods Enumerator values](https://senipah.github.io/VBA-Better-Array/api/SortMethods/SortMethods%20Enumeration.html). The default sort method is TimSort (`SM_TIMSORT`).

If the array has more than one dimension the `SortColumn` argument is used to determine the column in the array to be used for the comparison. `SortColumn` begins at base one, regardless of the bounds of the array at the second dimension.

#### Note

Arrays more than two dimensions deep are unsupported and an error will be raised when trying to sort them.

## Syntax

*expression*.**Sort**([*sortColumn*])

### Parameters

Name
: `sortColumn`

Type
: `Long`

Necessity
: Optional

Description
: The column in a two dimensional array to be used in the comparison.

### Returns

Type
: `BetterArray` / `Object`

Description
: The current instance of the BetterArray object with the array's order sorted in ascending order.

## Example

```vb
Public Sub SortExample()
    Dim result() As Variant
    Dim MyArray As BetterArray
    Set MyArray = New BetterArray

    MyArray.Push 3, 2, 6, 3, 7, 1, 3, 7, 9, 4
    MyArray.Sort
    result = MyArray.Items
    ' expected output:
    ' result =  1,2,3,3,3,4,6,7,7,9
End Sub
```

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)
