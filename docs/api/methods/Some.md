---
title: Some
parent: Methods
grand_parent: API
---

# Some

## Description
The `Some()` method determines whether at least one entry in the array matches `SearchElement`, returning `True` or `False` as appropriate.

`Some()` is provided as a JavaScript-style parity method and behaves the same as `Includes()`.

## Syntax

*expression*.**Some**(*SearchElement*, [*FromIndex*], [*Recurse*])

### Parameters

Name
: `SearchElement`

Type
: `Variant`

Necessity
: Required

Description
: The value to search for.

---

Name
: `FromIndex`

Type
: `Long`

Necessity
: Optional

Description
: The position in this array at which to begin searching for `SearchElement`; the first element to be searched is found at `FromIndex` for positive values of `FromIndex`, or at the array's `Length` property + `FromIndex` for negative values of `FromIndex`. Defaults to the array's `LowerBound` property.

---

Name
: `Recurse`

Type
: `Boolean`

Necessity
: Optional

Description
: If the array is jagged (an array of arrays) or multidimensional (stored internally as jagged arrays), set `Recurse` to `True` to search nested arrays.

### Returns

Type
: `Boolean`

Description
: `True` if at least one matching element is found; otherwise `False`.

## Example

```vb
Public Sub SomeExample()
    Dim MyArray As BetterArray
    Dim result As Boolean
    
    Set MyArray = New BetterArray
    MyArray.Push "Foo", "Bar", "Fizz", "Buzz"
    result = MyArray.Some("Bar")
    
    ' expected output:
    ' result is True
End Sub
```

## Inspiration
* <https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/some>

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)
