---
title: EveryType
parent: Methods
grand_parent: API
---

# EveryType

## Description
The `EveryType()` method determines whether all entries in the array are the same as the `SearchElement`, returning `True` or `False` as appropriate.

## Syntax

*expression*.**EveryType**(*SearchElement*, [*FromIndex*])

### Parameters

Name
: `SearchTypeName`

Type
: `String`

Necessity
: Required

Description
: The type name to search for.

---

Name
: `FromIndex`

Type
: `Long`

Necessity
: Optional

Description
: The position in this array at which to begin searching for `SearchElement`; the first character to be searched is found at `FromIndex` for positive values of `FromIndex`, or at the array's `Length` property + `FromIndex` for negative values of `FromIndex` (using the absolute value of `FromIndex` as the number of characters from the end of the array at which to start the search). Defaults to the array's `LowerBound` property.

### Returns

Type
: `Boolean`

Description
: `True` if the array includes `SearchElement`, `False` if not.

## Example

```vb
Public Sub EveryTypeExample()
    Dim MyArray As BetterArray
    Dim result As Boolean

    Set MyArray = New BetterArray
    MyArray.Push "Foo", "Foo", "Foo", "Foo"
    result = MyArray.EveryType("String")

    ' expected output:
    ' result is True
End Sub
```

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)
