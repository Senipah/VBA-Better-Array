---
title: IndexOf
parent: Methods
grand_parent: API
---

# IndexOf

## Description
The `IndexOf()` method method returns the first index at which a given element can be found in the array, or -1 if it is not present. 

## Syntax

*expression*.**IndexOf**(*SearchElement*,[*FromIndex*]) 

### Parameters

Name 
: `SearchElement`

Type
: `Variant`

Necessity
: Required

Description
: The element to locate in the array.

---

Name 
: `FromIndex`

Type
: `Long`

Necessity
: Optional

Description
: The index to start the search at. If the index is greater than or equal to the array's length, -1 is returned, which means the array will not be searched. If the provided index value is a negative number, it is taken as the offset from the end of the array. Note: if the provided index is negative, the array is still searched from front to back. If the provided index is 0, then the whole array will be searched. Default: 0 (entire array is searched).

### Returns

Type
: `Long`

Description
: The first index of the element in the array; -1 if not found.

## Inspiration
* <http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.indexof>
* <https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/indexOf>

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)