---
title: EveryType
parent: Methods
grand_parent: API
---

# EveryType

## Description
The `EveryType()` method determines whether all entries in the array are the same as the `searchElement`, returning `True` or `False` as appropriate.

## Syntax

*expression*.**EveryType**(*searchElement*, [*fromIndex*])

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
: The position in this array at which to begin searching for `searchElement`; the first character to be searched is found at `fromIndex` for positive values of `fromIndex`, or at the array's `Length` property + `fromIndex` for negative values of `fromIndex` (using the absolute value of `fromIndex` as the number of characters from the end of the array at which to start the search). Defaults to the array's `LowerBound` property.

### Returns

Type
: `Boolean` 

Description
: `True` if the array includes `searchElement`, `False` if not.


[Back to Docs](https://senipah.github.io/VBA-Better-Array/)
