---
title: IsSorted
parent: Methods
grand_parent: API
---

# IsSorted

## Description
The `IsSorted()` method tests if the stored array is sorted in ascending order. If a `ColumnIndex` argument is provided and the array is jagged or multi-dimensional, it will test if the aray is sorted by the values in that column. 

#### Note

`IsSorted` will raise an error if the array is more than two dimensions deep.

## Syntax

*expression*.**IsSorted**(`ColumnIndex`) 

### Parameters

Name 
: `ColumnIndex`

Type
: `Long`

Necessity
: Optional

Description
: The index of the column to be extracted. Only applies to multi-dimensional or jagged arays. If ommitted, or exceeds the bounds of the array, the first column will be returned.

### Returns

Type
: `Boolean`

Description
: `True` if the array is sorted, `False` if not.

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)