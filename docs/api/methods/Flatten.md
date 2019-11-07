---
title: Flatten
parent: Methods
grand_parent: API
---

# Flatten

## Description
The `Flatten()` method flattens a stored multi-dimensional or jagged array into a single one-dimensional array. Returns the current instance with it's stored array flattened.

## Syntax

*expression*.**Flatten**()

### Parameters

**None**

### Returns

Type
: `BetterArray` / `Object`

Description
: The current instance with it's stored array flattened.

## Example

```vb
Public Sub FlattenExample()
    Dim MultiDimensionArray(1 To 2, 1 To 2) As Variant
    Dim MyArray As BetterArray
    Dim result() As Variant
    Set MyArray = New BetterArray
    
    MultiDimensionArray(1, 1) = "Foo"
    MultiDimensionArray(1, 2) = "Bar"
    MultiDimensionArray(2, 1) = "Fizz"
    MultiDimensionArray(2, 2) = "Buzz"
    
    MyArray.Items = MultiDimensionArray
    MyArray.Flatten

    result = MyArray.Items

    ' expected output:
    ' result is a one-dimension array with the values: "Foo", "Bar", "Fizz", "Buzz"
End Sub
```

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)