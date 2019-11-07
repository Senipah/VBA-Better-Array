---
title: Unique
parent: Methods
grand_parent: API
---

# Unique

## Description
The `Unique()` method removes all duplicate elements from the outermost array such that only unique elements remain.

## Syntax

*expression*.**Unique**() 

### Parameters

**None**

### Returns

Type
: `BetterArray` / `Object`

Description
: The current instance of the BetterArray object with any duplicate elements removed.

## Example

```vb
Public Sub UniqueExample()
    Dim result() As Variant
    Dim MyArray As BetterArray
    Set MyArray = New BetterArray
    
    MyArray.Push "Foo", "Foo", "Bar", "Foo", "Buzz", "Bar"
    MyArray.Unique
    result = MyArray.Items
    ' expected output:
    ' result is an array containing "Foo","Bar","Buzz"
End Sub
```

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)