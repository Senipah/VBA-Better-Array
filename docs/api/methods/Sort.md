---
title: Sort
parent: Methods
grand_parent: API
---

# Sort

## Description
The `Sort()` method sorts and returns the stored array. If the array has more than one dimension the `SortColumn` argument is used to determine the column in the array to be used for the comparison. Arrays with more than two dimensions are unsupported and will be returned unchanged. Uses an implementation of the [Quicksort](https://en.wikipedia.org/wiki/Quicksort) algorithm.

## Syntax

*expression*.**Sort**([*SortColumn*])

### Parameters

Name 
: `SortColumn`

Type
: `Long`

Necessity
: Optional

Description
: The column in a two dimensional array to be used in the comparison.

### Returns

Type
: BetterArray `Object`

Description
: The current instance of the BetterArray object with the array's order sorted in ascending order.

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)