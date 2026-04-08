---
title: LastIndexOf
parent: Methods
grand_parent: API
---

# LastIndexOf

## Description
The `LastIndexOf()` method method returns the last index at which a given element can be found in the array, or -9999 if it is not present.

## Syntax

*expression*.**LastIndexOf**(*SearchElement*,[*FromIndex*],[*CompType*])

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
: The index at which to begin searching backwards. If omitted, search starts at the array's `UpperBound`. If a value above `UpperBound` is provided, it is clamped to `UpperBound`. If the provided index is lower than `LowerBound` and negative, it is treated as an offset from the end (`LowerBound + Length + FromIndex`) and then clamped to `LowerBound` if needed.

---

Name
: `CompType`

Type
: `ComparisonType (Long)`

Necessity
: Optional

Description
: The type of comparison to perform. See the [comparison type enumeration](https://senipah.github.io/VBA-Better-Array/api/enumerations/ComparisonType_Enumeration.html). Default is CT_EQUALITY.

#### ComparisonType Enumerations

`CT_EQUALITY`
: Compares `SearchElement` against the element at the current index for equality. This is the default comparison method.

`CT_LIKENESS`
: `SearchElement` is treated as a string pattern and compared against the element as the current index using the `Like` operator. If this option is chosen `SearchElement` must be a String type or an error will be raised.


### Returns

Type
: `Long`

Description
: The last index of the element in the array; -9999 if not found.

## Example

```vb
Public Sub LastIndexOfExample()
    Dim result As Long
    Dim MyArray As BetterArray
    Set MyArray = New BetterArray

    MyArray.Push "Apple", "Banana", "Orange", "Apple", "Mango"
    result = MyArray.LastIndexOf("Apple")
    ' expected output:
    ' result  = 3
End Sub
```


## Inspiration
* <http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.lastindexof>
* <https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/lastIndexOf>

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)
