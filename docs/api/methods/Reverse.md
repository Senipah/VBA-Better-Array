---
title: Reverse
parent: Methods
grand_parent: API
---

# Reverse

## Description
The `Reverse()` method reverses the order of elements in the array. The first array element becomes the last, and the last array element becomes the first.

## Syntax

*expression*.**Reverse**([*Recurse*])

### Parameters

Name
: `Recurse`

Type
: `Boolean`

Necessity
: Optional

Description
: If true and the array is jagged or multi-dimensional (which are stored as jagged internally), the order of all nested arrays will also be reversed.

### Returns

Type
: `BetterArray` / `Object`

Description
: The current instance of the BetterArray object with the array's order reversed.

## Example

```vb
Public Sub ReverseExample()
    Dim result() As Variant
    Dim MyArray As BetterArray
    Set MyArray = New BetterArray

    MyArray.Push "Banana", "Orange", "Apple", "Mango"
    MyArray.Reverse
    result = MyArray.Items
    ' expected output:
    ' result = "Mango","Apple","Orange","Banana"
End Sub
```

## Inspiration
* <http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.reverse>
* <https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/reverse>

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)
