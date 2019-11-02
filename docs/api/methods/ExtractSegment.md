---
title: ExtractSegment
parent: Methods
grand_parent: API
---

# ExtractSegment

## Description
The `ExtractSegment()` method extracts the specified segment of an array. If the current instance stores a two-dimensional array, you can enter a row or column index to return the specified segment of the array as a one-dimension array. If both row and column index arguments are provided the element stored at the intersection will be returned (wrapped in an array if the element is not already an array). 

If the stored array is one-dimension and both column and row arguments are provided, the element at the row index will be returned (encased in an array). If just a row or just a column index are provided the element at whichever index has been provided will be returned (encased in an array).

## Syntax

*expression*.**ExtractSegment**([*rowIndex*], [*columnIndex*])

### Parameters

Name 
: `rowIndex`

Type
: `Long`

Necessity
: Optional

Description
: The index of the row to be extracted. 

---

Name 
: `columnIndex`

Type
: `Long`

Necessity
: Optional

Description
: The index of the column to be extracted. 

### Returns

Type
: `Variant()`

Description
: A variant array containing the extracted segment.


[Back to Docs](https://senipah.github.io/VBA-Better-Array/)