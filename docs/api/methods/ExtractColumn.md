---
title: ExtractColumn
parent: Methods
grand_parent: API
---

# ExtractColumn

## Description
The `ExtractColumn()` method extracts the specified column of a multidimensional/jagged array and returns it as a new array. This does not mutate the stored array. If operating on a one-dimension array, a copy of the stored array will be returned.

## Syntax

*expression*.**ExtractColumn**([*ColumnNumber*])

### Parameters

Name 
: `ColumnNumber`

Type
: `Long`

Necessity
: Optional

Description
: The base-1 index of the column to be extracted. If ommitted, or exceeds the bounds of the array, the first column will be returned. 

### Returns

Type
: `Variant()`

Description
: A variant array containing the extracted column.

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)