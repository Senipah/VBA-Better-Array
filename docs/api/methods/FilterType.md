---
title: FilterType
parent: Methods
grand_parent: API
---


# FilterType

## Description
The `FilterType()` method filters and returns the current array based on the specified filter criteria. 

#### Note
Provides ability to filter on any type of array (unlike the VBA version which only works with `String` arrays).

## Syntax

*expression*.**FilterType**(*match*, [*exclude*], [*recurse*])

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
: `exclude`

Type
: `Boolean`

Necessity
: Optional

Description
: Boolean value indicating whether to return values that include or exclude `match`. If include is True, `FilterType` returns the subset of the array that contains `match`. If include is False, `FilterType` returns the subset of the array that does not contain `match`.

---

Name 
: `recurse`

Type
: `Boolean`

Necessity
: Optional

Description
: Boolean value indicating whether the filter should be applied recrsively to a jagged or multidimensional array. By default, only the outermost array will be filtered.

### Returns

Type
: BetterArray `Object`

Description
: The current instance of the BetterArray object with the filter applied to the stored array. 

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)