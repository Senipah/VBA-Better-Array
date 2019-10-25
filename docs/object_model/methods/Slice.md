---
title: Slice
parent: Methods
grand_parent: API
---

# Slice

## Description
The `Slice()` method returns a shallow copy of a portion of an array into a new array object selected from begin to end (end not included) where begin and end represent the index of items in that array. The original array will not be modified.

## Arguments
### `StartIndex` (Long)
The index at which to begin extraction.
A negative index can be used, indicating an offset from the end of the sequence. slice(-2) extracts the last two elements in the sequence.
If begin is undefined, slice begins from index 0.
If begin is greater than the length of the sequence, an empty array is returned.
### *Optional* `EndIndex` (Long)
The index before which to end extraction. slice extracts up to but not including end.
For example, slice(1,4) extracts the second element through the fourth element (elements indexed 1, 2, and 3).
A negative index can be used, indicating an offset from the end of the sequence. slice(2,-1) extracts the third element through the second-to-last element in the sequence.
If end is omitted, slice extracts through the end of the sequence (arr.length).
If end is greater than the length of the sequence, slice extracts through to the end of the sequence (arr.length).

## Returns
### (Variant)
A new array containing the extracted elements.

## Inspiration
* <http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.slice>
* <https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/slice>


[Back to Docs](https://senipah.github.io/VBA-Better-Array/)
