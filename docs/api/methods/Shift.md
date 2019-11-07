---
title: Shift
parent: Methods
grand_parent: API
---

# Shift

## Description
The `Shift()` method removes the first element from the array and returns that removed element. This method changes the length of the array.

## Syntax

*expression*.**Shift**()

### Parameters

**None**

### Returns

Type
: `Variant`

Description
: The first element from the array.

## Example

```vb
Public Sub ShiftExample()
    Dim result As String
    Dim MyArray As BetterArray
    Set MyArray = New BetterArray
    
    MyArray.Push "Banana", "Orange", "Apple", "Mango"
    result = MyArray.Shift
    ' expected output:
    ' result = "Banana"
    ' MyArray contains: "Orange", "Apple", "Mango"
End Sub
```

## Inspiration
* <https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/shift>
* <http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.shift>

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)