---
title: Transpose
parent: Methods
grand_parent: API
---

# Transpose

## Description

The `Transpose()` method converts rows to columns and vice versa.

## Syntax

*expression*.**Transpose**()

### Parameters

**None**

### Returns

Type
: `BetterArray` / `Object`

Description
: The current instance with it's stored array transposed.

## Example

```vb
Public Sub TransposeExample()
    Dim result() As Variant
    Dim MyArray As BetterArray
    Set MyArray = New BetterArray
    
    MyArray.Push Array("Foo", 1)
    MyArray.Push Array("Bar", 2)
    MyArray.Transpose
    result = MyArray.Items
    ' expected output:
    ' result is a jagged array with the structure "[["Foo","Bar"],[1,2]]"
End Sub
```


[Back to Docs](https://senipah.github.io/VBA-Better-Array/)





