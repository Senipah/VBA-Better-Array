---
title: Concat
parent: Methods
grand_parent: API
---

# Concat

## Description
The `Concat()` method joins one or more arrays onto the end of the current array.

## Syntax

*expression*.**Concat**([*args1*[, *args2*[, ...[, *argsN*]]]])

### Parameters

Name 
: `args`

Type
: ParamArray `Variant`

Necessity
: Optional

Description
: The array(s) to be added to the end of the array. 

### Returns

Type
: `BetterArray` / `Object`

Description
: The current instance of the BetterArray object with the passed arrays having been added to the end. 

## Example

```vb
Public Sub ConcatExample()
    Dim firstItems() As Variant
    Dim secondItems() As Variant
    Dim result() As Variant
    Dim MyArray As BetterArray
    
    Set MyArray = New BetterArray
    firstItems = Array("Foo", "Bar")
    secondItems = Array("Fizz", "Buzz")
    
    MyArray.Items = firstItems
    MyArray.Concat secondItems
    
    result = MyArray.Items
    
    ' expected output:
    ' result is an array with the values: "Foo", "Bar", "Fizz", "Buzz"
End Sub
```



## Inspiration
* <https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/concat>
* <http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.concat>

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)