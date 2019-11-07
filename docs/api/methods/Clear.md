---
title: Clear
parent: Methods
grand_parent: API
---

# Clear

## Description
The `Clear()` method clears all entries in the current array but retains the same internal capacity.

## Syntax

*expression*.**Clear**()

### Parameters

**None**

### Returns

Type
: `BetterArray` / `Object`

Description
: The current instance of the BetterArray object with all entries cleared but capacity retained. 

## Example

```vb
Public Sub ClearExample()
    Dim result() As Variant
    Dim MyArray As BetterArray
    Set MyArray = New BetterArray
    MyArray.Push 1, 2, 3
    MyArray.Clear
    result = MyArray.Items
    ' expected output: result is an empty array
End Sub
```

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)