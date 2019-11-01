---
title: CopyWithin
parent: Methods
grand_parent: API
---

# CopyWithin

## Description
The `CopyWithin()` method copies part of an array to another location in the same array without modifying its length. The `CopyWithin` method takes up to three Parameters `Target`, `StartIndex` and `EndIndex`.

#### Note
The end argument is optional with the length of the this object as its default value. If target is negative, it is treated as length + target where length is the length of the array. If start is negative, it is treated as length + start. If end is negative, it is treated as length + end.


## Syntax

*expression*.**CopyWithin**(*target*, [*startIndex*], [*endIndex*])

### Parameters

Name 
: `target`

Type
: `Long`

Necessity
: Required

Description
: The index at which to copy the sequence to. If negative, `target` will be counted from the end. If `target` is at or greater than the array's `Length` property, nothing will be copied. If `target` is positioned after `startIndex`, the copied sequence will be trimmed to fit the array's `Length` property.

---

Name
: `startIndex`

Type
: `Long`

Necessity
: Optional

Description
: The index at which to start copying elements from. If negative, `startIndex` will be counted from the end. If `startIndex` is omitted, `CopyWithin` will copy from the LowerBound index of the array. 

---

Name
: `endIndex`

Type
: `Long`

Necessity
: Optional

Description
: The index at which to end copying elements from. `CopyWithin` copies up to but not including `endIndex`. If negative, `endIndex` will be counted from the end.
If `endIndex` is omitted, `CopyWithin` will copy until the last index (default to the array's `Length` property).

### Returns

Type
: `BetterArray` / `Object`

Description
: The current instance of the BetterArray object with the stored array modified. 


## Inspiration
* <http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.copywithin>
* <https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/copyWithin>


[Back to Docs](https://senipah.github.io/VBA-Better-Array/)