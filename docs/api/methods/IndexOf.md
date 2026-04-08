---
title: IndexOf
parent: Methods
grand_parent: API
---

# IndexOf

## Description
The `IndexOf()` method method returns the first index at which a given element can be found in the array, or -9999 if it is not present.

## Syntax

*expression*.**IndexOf**(*SearchElement*,[*FromIndex*],[*CompType*])

### Parameters

Name
: `SearchElement`

Type
: `Variant`

Necessity
: Required

Description
: The element to locate in the array.

---

Name
: `FromIndex`

Type
: `Long`

Necessity
: Optional

Description
: The index at which to begin searching. If omitted, search starts at the array's `LowerBound`. If the provided index is lower than `LowerBound` and negative, it is treated as an offset from the end (`LowerBound + Length + FromIndex`) and then clamped to `LowerBound` if needed. If the resulting start index is greater than `UpperBound`, `-9999` is returned.

---

Name
: `CompType`

Type
: `ComparisonType (Long)`

Necessity
: Optional

Description
: The type of comparison to perform. See the [comparison type enumeration](https://senipah.github.io/VBA-Better-Array/api/enumerations/ComparisonType_Enumeration.html). Default is CT_EQUALITY.

### Returns

Type
: `Long`

Description
: The first index of the element in the array; -9999 if not found.

## Example

```vb
Public Sub IndexOfExample()
    Dim result As Long
    Dim MyArray As BetterArray
    Set MyArray = New BetterArray

    MyArray.Push "Banana", "Orange", "Apple", "Mango"
    result = MyArray.IndexOf("Apple")
    ' expected output:
    ' result = 2
End Sub
```

## Inspiration
* <http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.indexof>
* <https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/indexOf>

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)
