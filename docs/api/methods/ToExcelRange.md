---
title: ToExcelRange
parent: Methods
grand_parent: API
---

# ToExcelRange

## Description

The `ToExcelRange()` method writes the values stored in the array to an excel worksheet starting at the specified range."


## Syntax

*expression*.**ToExcelRange**(*destination*, [*transposeValues*])

### Parameters

Name 
: `destination`

Type
: `Range` / `Object`

Necessity
: Required

Description
: An Excel Range object representing the loction to begin writing the stored values. Will be expanded as necessar to accomodate the size of the array.

---

Name 
: `transposeValues`

Type
: `Boolean`

Necessity
: Optional

Description
: If present, stored values will be transposed when written to the Excel Range (rows become columns and vice versa)

### Returns

Type
: `Range` / `Object`

Description
: The Excel Range object containing the outputted values.


[Back to Docs](https://senipah.github.io/VBA-Better-Array/)





