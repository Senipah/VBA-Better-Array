---
title: Fill
parent: Methods
grand_parent: API
---


# Fill

## Description
The `Fill()` method method fills (modifies) all the elements of an array from a start index (default zero) to an end index (default array length) with a static value.

## Syntax

*expression*.**Fill**(*value*, [*startIndex*], [*endIndex*])

### Parameters

Name 
: `value`

Type
: `Variant`

Necessity
: Required

Description
: The value to populate the array with.

---

Name
: `startIndex`

Type
: `Long`

Necessity
: Optional

Description
: The first index of the outermost array to begin filling with the passed value. If ommited the array will be filled from the first index.

---

Name 
: `endIndex`

Type
: `Boolean`

Necessity
: Optional

Description
: The last index of the outermost array to to fill up to with the passed value.

### Returns

Type
: BetterArray `Object`

Description
: The current instance of the BetterArray object array filled with the passed value between the specified indices. If ommited the array will be filled to the last index.

## Inspiration
* <http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.fill>
* <https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/fill>

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)