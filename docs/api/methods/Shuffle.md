---
title: Shuffle
parent: Methods
grand_parent: API
---

# Shuffle

## Description
The `Shuffle()` method Shuffles the order of elements in the array using the [Fisher–Yates algorithm](https://en.wikipedia.org/wiki/Fisher%E2%80%93Yates_shuffle#The_modern_algorithm).

## Syntax

*expression*.**Shuffle**([*recurse*])

### Parameters

Name
: `recurse`

Type
: `Boolean`

Necessity
: Optional

Description
: If true and the array is jagged or multi-dimensional (which are stored as jagged internally), the order of all nested arrays will also be Shuffled.

### Returns

Type
: `BetterArray` / `Object`

Description
: The current instance of the BetterArray object with the array's order Shuffled.


## Example

```vb
Public Sub ShiftExample()
    Dim result() As Variant
    Dim MyArray As BetterArray
    Set MyArray = New BetterArray

    MyArray.Push "Banana", "Orange", "Apple", "Mango"
    MyArray.Shuffle
    result = MyArray.Items
    ' expected output:
    ' result will contain  "Banana", "Orange", "Apple", "Mango" but the order
    ' has been shuffled. e.g: "Mango", "Banana", "Apple", "Orange"
End Sub
```

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)
