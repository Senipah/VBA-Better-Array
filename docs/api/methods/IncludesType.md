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

*expression*.**IncludesType**(*searchTypeName*, [*fromIndex*])

### Parameters

Name 
: `searchTypeName`

Type
: `String`

Necessity
: Required

Description
: The type name to search for.

---

Name 
: `fromIndex`

Type
: `Long`

Necessity
: Optional

Description
: The position in this array at which to begin searching for `searchTypeName`; the first character to be searched is found at `fromIndex` for positive values of `fromIndex`, or at the array's `Length` property + `fromIndex` for negative values of `fromIndex` (using the absolute value of `fromIndex` as the number of characters from the end of the array at which to start the search). Defaults to the array's `LowerBound` property.

### Returns

Type
: `Boolean`

Description
: `True` if the array includes `searchTypeName` type, `False` if not.

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)
