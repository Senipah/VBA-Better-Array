---
title: FromExcelRange
parent: Methods
grand_parent: API
---

# FromExcelRange

## Description

The `FromExcelRange()` method accepts an Excel Range object as an argument and stores the values contained within the range to the internal array. If the Range object contains only one row or column, the values will be stored and returned as a one-dimension array. 

If the Range contains both multiple Rows and Columns the default behaviour is that the array returned by accessing the `Items` get accessor will be returned as a multi-dimension array. If a `columnNumber` or `rowNumber` argument is provided, only the values in that Row/Column will be stored as a one-dimension array.


## Syntax

*expression*.**FromExcelRange**(*fromRange*, [*columnNumber*], [*rowNumber*])

### Parameters

Name 
: `fromRange`

Type
: `Range` / `Object`

Necessity
: Required

Description
: An Excel Range object containing the values to be stored in the array.

---

Name 
: `columnNumber`

Type
: `Long`

Necessity
: Optional

Description
: If present, only the values in this column of the range will be stored.

---

Name 
: `rowNumber`

Type
: `Long`

Necessity
: Optional

Description
: The current instance of the BetterArray object with the specified range of values stored in the internal array.

### Returns

Type
: `BetterArray` / `Object`

Description
: If present, only the values in this row of the range will be stored.


[Back to Docs](https://senipah.github.io/VBA-Better-Array/)





