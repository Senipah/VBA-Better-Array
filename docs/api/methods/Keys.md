---
title: Keys
parent: Methods
grand_parent: API
---

# Keys

## Description
The `Keys()` method returns a new array (0 based) that contains the keys for each index in the array.

## Syntax

*expression*.**Keys**()

### Parameters

**None**

### Returns

Type
: `Variant()`

Description
: A 0-based `Variant` array containing the indexes used in the outermost internal array.


## Example

```vb
Public Sub KeysExample()
    Dim result() As Variant
    Dim MyArray As BetterArray
    Set MyArray = New BetterArray
    
    MyArray.LowerBound = 10
    MyArray.Push 1, 2, 3, 4, 5, 6
    result = MyArray.Keys
    ' expected output:
    ' result is a zero-based array with the values: 10, 11, 12, 13, 14, 15
End Sub
```



## Inspiration
* <http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.keys>
* <https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/keys>


[Back to Docs](https://senipah.github.io/VBA-Better-Array/)
