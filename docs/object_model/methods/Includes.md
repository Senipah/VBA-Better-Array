---
title: Includes
parent: Methods
grand_parent: API
---

# Includes

## Description
The `Includes()` method determines whether the array includes a certain value among its entries, returning `True` or `False` as appropriate.

## Arguments
### `SearchElement` (Variant)
The value to search for.
### *Optional* `FromIndex` (Long)
The position in this array at which to begin searching for `SearchElement`; the first character to be searched is found at `FromIndex` for positive values of `FromIndex`, or at the array's `Length` property + `FromIndex` for negative values of `FromIndex` (using the absolute value of `FromIndex` as the number of characters from the end of the array at which to start the search). Defaults to the array's `LowerBound` property.
## Returns
### (Boolean)
`True` if the array includes valueToFind, `False` if not.

## Inspiration
* <http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.includes>
* <https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/includes>


[Back to Docs](https://senipah.github.io/VBA-Better-Array/)
