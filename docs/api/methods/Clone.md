---
title: Clone
parent: Methods
grand_parent: API
---

# Clone

## Description
The `Clone()` method returns a new BetterArray instance containing the same values as the current instance.

## Syntax

*expression*.**Clone**()

### Parameters

**None**

### Returns

Type
: `BetterArray` / `Object`

Description
: A BetterArray instance containing the same values as the current instance.

## Example

```vb
Public Sub CloneExample()
    Dim First As BetterArray
    Dim Second As BetterArray
    
    Set First = New BetterArray
    First.Push 1, 2, 3
    Set Second = First.Clone
    First.Clear
    
    Dim firstContents() As Variant
    Dim secondContents() As Variant
    firstContents = First.Items
    secondContents = Second.Items
    
    ' expected output:
    ' firstContents is an empty array
    ' secondContents is an array with the values: 1,2,3
End Sub
```

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)