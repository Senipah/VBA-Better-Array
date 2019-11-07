---
title: ResetToDefault
parent: Methods
grand_parent: API
---

# ResetToDefault

## Description
The `ResetToDefault()` method clears all entries in the current array and resets capacity to default value.

## Syntax

*expression*.**ResetToDefault**()

### Parameters

**None**

### Returns

Type
: `BetterArray` / `Object`

Description
: The current instance of the BetterArray object with all entries cleared and the capacity reset. 

## Example

```vb
Public Sub ResetToDefaultExample()
    Dim Capacity As Long
    Dim MyArray As BetterArray
    Set MyArray = New BetterArray
    
    MyArray.Push 1, 2, 3, 4, 5, 6, 7, 8, 9, 10
    MyArray.ResetToDefault
    Capacity = MyArray.Capacity
    ' expected output:
    ' Capacity = 4 (the default capacity)
    ' MyArray contains no entries
End Sub

```

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)