---
title: IncludesType
parent: Methods
grand_parent: API
---

# IncludesType

## Description
The `IncludesType()` method determines whether the array includes a certain type among its entries, returning `True` or `False` as appropriate.

#### Note

A list of data types supported in VBA is available in the official language documentation [here](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/data-type-summary).

## Syntax

*expression*.**IncludesType**(*SearchTypeName*, [*FromIndex*])

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
: The position in this array at which to begin searching for `SearchTypeName`; the first character to be searched is found at `FromIndex` for positive values of `FromIndex`, or at the array's `Length` property + `FromIndex` for negative values of `FromIndex` (using the absolute value of `FromIndex` as the number of characters from the end of the array at which to start the search). Defaults to the array's `LowerBound` property.

---

Name
: `recurse`

Type
: `Boolean`

Necessity
: Optional

Description
: If the array is jagged (an array of arrays) or multidimensional (which are stored internally as jagged arays) and you wish for all nested arrays to be checked then `recurse` must be true - otherwise only the outermost array will be checked. This argument has no effect when operating on a one-dimension array.


### Returns

Type
: `Boolean`

Description
: `True` if the array includes `SearchTypeName` type, `False` if not.

## Example

```vb
Public Sub IncludesTypeExample()
    Dim result As Boolean
    Dim MyArray As BetterArray
    Set MyArray = New BetterArray

    MyArray.Push "Foo", 1.23, "Fizz", "Buzz"
    result = MyArray.IncludesType("Double")
    ' expected output:
    ' result is True
End Sub
```



[Back to Docs](https://senipah.github.io/VBA-Better-Array/)
