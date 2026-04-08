---
title: At
parent: Methods
grand_parent: API
---

# At

## Description
The `At()` method returns the element at a relative index.

Unlike the default `Item` property, `At()` uses zero-based relative positions:
- `0` returns the first element,
- `1` returns the second element,
- `-1` returns the last element.

If the array is empty, or if the requested relative index is out of range, `At()` returns `Empty`.

## Syntax

*expression*.**At**(*RelativeIndex*)

### Parameters

Name
: `RelativeIndex`

Type
: `Long`

Necessity
: Required

Description
: The relative position to retrieve.
Use non-negative values to count from the start of the array, and negative values to count from the end.

### Returns

Type
: `Variant`

Description
: The element at the requested relative index, or `Empty` if out of range.

## Example

```vb
Public Sub AtExample()
    Dim MyArray As BetterArray
    Dim firstValue As Variant
    Dim lastValue As Variant
    
    Set MyArray = New BetterArray
    MyArray.Push "Banana", "Orange", "Apple", "Mango"
    
    firstValue = MyArray.At(0)
    lastValue = MyArray.At(-1)
    
    ' expected output:
    ' firstValue = "Banana"
    ' lastValue = "Mango"
End Sub
```

## Inspiration
* <https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/at>

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)
