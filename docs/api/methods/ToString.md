---
title: ToString
parent: Methods
grand_parent: API
---

# ToString

## Description
The `ToString()` method returns a string representing the array structure and its elements. Arrays are enclosed with square brackets. Elements are comma-separated. Object types are represented as "[Object]". Set the `PrettyPrint` argument to `True` to format the returned string with indentation for easier viewing of long or nested arrays.

## Syntax

*expression*.**ToString**([*PrettyPrint*], [*DelimitStrings*], [*OpeningDelimiter*], [*ClosingDelimiter*])

### Parameters

Name 
: `PrettyPrint`

Type
: `Boolean`

Necessity
: Optional

Description
: Set to true to format the returned string for easier viewing.

---

Name 
: `DelimitStrings`

Type
: `Boolean`

Necessity
: Optional

Description
: If true, any string values stored in the array will additionally be enclosed by opening and closing quotation marks (`"`).

---

Name 
: `OpeningDelimiter`

Type
: `String`

Necessity
: Optional

Description
: If provided, the string passed to `OpeningDelimiter` will be used to mark the beginning of arrays. If ommitted, the default character of `{` will be used.

---

Name 
: `ClosingDelimiter`

Type
: `String`

Necessity
: Optional

Description
: If provided, the string passed to `ClosingDelimiter` will be used to mark the beginning of arrays. If ommitted, the default character of `}` will be used.

### Returns

Type
: `String`

Description
: A string representing the array structure and its elements.


[Back to Docs](https://senipah.github.io/VBA-Better-Array/)