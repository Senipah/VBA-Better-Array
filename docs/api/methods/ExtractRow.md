---
title: ExtractRow
parent: Methods
grand_parent: API
---

# ExtractRow

## Description
The `ExtractRow()` method extracts the specified row of a multidimensional/jagged array and returns it as a new array. This does not mutate the stored array. If operating on a one-dimension array, a copy of the stored array will be returned.

## Syntax

*expression*.**ExtractRow**([*RowNumber*])

### Parameters

Name 
: `RowNumber`

Type
: `Long`

Necessity
: Optional

Description
: The base-1 index of the row to be extracted. If ommitted, or exceeds the bounds of the array, the first row will be returned. 

### Returns

Type
: `Variant()`

Description
: A variant array containing the extracted row.

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)