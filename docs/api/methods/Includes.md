---
title: Includes
parent: Methods
grand_parent: API
---

# Includes

## Description
The `Includes()` method determines whether the array includes a certain value among its entries, returning `True` or `False` as appropriate.

## Syntax

*expression*.**Includes**(*searchElement*, [*fromIndex*])

### Parameters

Name 
: `searchElement`

Type
: `Variant`

Necessity
: Required

Description
: The value to search for.

---

Name 
: `fromIndex`

Type
: `Long`

Necessity
: Optional

Description
: The position in this array at which to begin searching for `searchElement`; the first character to be searched is found at `fromIndex` for positive values of `fromIndex`, or at the array's `Length` property + `fromIndex` for negative values of `fromIndex` (using the absolute value of `fromIndex` as the number of characters from the end of the array at which to start the search). Defaults to the array's `LowerBound` property.

---

Name 
: `recurse`

Type
: `Boolean`

Necessity
: Optional

Description
: If the array is jagged (an array of arrays) or multidimensional (which are stored internally as jagged arays) and you wish for all nested arrays to be checked then `recurse` must be true - otherwise only the outermost array will be checked. This argument has no effect when operating on a one-dimension array.


### Returns

Type
: `Boolean`

Description
: `True` if the array includes `searchElement`, `False` if not.

## Inspiration
* <http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.includes>
* <https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/includes>


[Back to Docs](https://senipah.github.io/VBA-Better-Array/)
