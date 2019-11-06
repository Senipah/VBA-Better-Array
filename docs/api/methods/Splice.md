---
title: Splice
parent: Methods
grand_parent: API
---

# Splice

## Description
The `Splice()` method returns a shallow copy of a portion of an array into a new array selected from begin to end (end not included) where begin and end represent the index of items in that array. The original array will not be modified.

## Syntax

*expression*.**Splice**(*startIndex*, [*endIndex*])

### Parameters

Name 
: `startIndex`

Type
: `Long`

Necessity
: Required

Description
: The index at which to begin extraction. A negative index can be used, indicating an offset from the end of the sequence. Splice(-2) extracts the last two elements in the sequence. If begin is undefined, Splice begins from the value of the `LowerBound` property. If begin is greater than the length of the sequence, an empty array is returned.

---

Name 
: `endIndex`

Type
: `Long`

Necessity
: Optional

Description
: The index before which to end extraction. Splice extracts up to but not including end.
For example, `Splice(1,4)` extracts the element at index 1 through the element indexed at 4 (elements indexed 1, 2, and 3). A negative index can be used, indicating an offset from the end of the sequence. `Splice(2,-1)` extracts the third element through the second-to-last element in the sequence. If `endIndex` is omitted or is greater than the `UpperBound` value, Splice extracts through the end of the sequence (the `UpperBound` value).

### Returns

Type
: `Variant()`

Description
: A new array containing the extracted elements.

## Inspiration
* <http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.Splice>
* <https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/Splice>


[Back to Docs](https://senipah.github.io/VBA-Better-Array/)
