---
title: Remove
parent: Methods
grand_parent: API
---

# Remove

## Description
The `Remove()` method removes the element at the specified index from the array. If the index is negative, the index will be counted from the end of the array. Returns the new length of the array.

## Syntax

*expression*.**Remove**(*index*)

### Parameters

Name
: `Index`

Type
: `Long`

Necessity
: Required

Description
: The index of the the element to be removed from the array.

### Returns

Type
: `Long`

Description
: The new length of the array.

## Example

```vb
Public Sub RemoveExample()
    Dim result As Long
    Dim MyArray As BetterArray
    Set MyArray = New BetterArray

    MyArray.Push "Banana", "Orange", "Apple", "Mango"
    result = MyArray.Remove(2)

    ' expected output:
    ' result  = 3 - the new length of MyArray
    ' MyArray now contains "Banana", "Orange", "Mango"
End Sub
```

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)
