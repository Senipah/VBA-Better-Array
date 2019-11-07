---
title: Pop
parent: Methods
grand_parent: API
---

# Pop

## Description
The `Pop()` method removes the last element from the array and returns that element. This method changes the length of the array.

## Syntax

*expression*.**Pop**()

### Parameters

**None**

### Returns

Type
: `Variant`

Description
: The last element from the array.


## Example

```vb
Public Sub PopExample()
    Dim result As String
    Dim MyArray As BetterArray
    Set MyArray = New BetterArray
    
    MyArray.Push "Banana", "Orange", "Apple", "Mango"
    result = MyArray.Pop
    ' expected output:
    ' result  = "Mango"
    ' MyArray now contains "Banana", "Orange", "Apple"
End Sub
```

## Inspiration
* <https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/pop>
* <http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.pop>


[Back to Docs](https://senipah.github.io/VBA-Better-Array/)