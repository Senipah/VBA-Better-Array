---
title: Unshift
parent: Methods
grand_parent: API
---

# Unshift

## Description
The `Unshift()` method adds one or more elements to the beginning of an array and returns the new length of the array.

## Syntax

*expression*.**Unshift**([*args1*[, *args2*[, ...[, *argsN*]]]])

### Parameters

Name 
: `args`

Type
: ParamArray `Variant`

Necessity
: Optional

Description
: The element(s) to be added to the beginning of the array

### Returns

Type
: `Long`

Description
: The new length of the array.

## Example

```vb
Public Sub UnshiftExample()
    Dim result() As Variant
    Dim MyArray As BetterArray
    Set MyArray = New BetterArray
    
    MyArray.Push "Banana", "Orange", "Apple", "Mango"
    MyArray.Unshift "Lemon", "Pineapple"
    result = MyArray.Items
    ' expected output:
    ' result is an array containing: "Lemon", "Pineapple","Banana", "Orange", "Apple", "Mango"
End Sub
```

## Inspiration
* <https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/Unshift>
* <http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.unshift>

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)