---
title: LastIndexOf
parent: Methods
grand_parent: API
---

# IndexOf

## Description
The `LastIndexOf()` method method returns the last index at which a given element can be found in the array, or -9999 if it is not present. 

## Syntax

*expression*.**LastIndexOf**(*searchElement*,[*fromIndex*]) 

### Parameters

Name 
: `searchElement`

Type
: `Variant`

Necessity
: Required

Description
: The element to locate in the array.

---

Name 
: `fromIndex`

Type
: `Long`

Necessity
: Optional

Description
: The index to start the search at. If the index is greater than or equal to the array's length, -9999 is returned, which means the array will not be searched. If the provided index value is a negative number, it is taken as the offset from the end of the array. Note: if the provided index is negative, the array is still searched from front to back. If the provided index is 0, then the whole array will be searched. Default: entire array is searched.

### Returns

Type
: `Long`

Description
: The last index of the element in the array; -9999 if not found.

## Inspiration
* <http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.lastindexof>
* <https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/lastIndexOf>

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)