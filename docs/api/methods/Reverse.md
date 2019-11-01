---
title: Reverse
parent: Methods
grand_parent: API
---

# Reverse

## Description
The `Reverse()` method reverses the order of elements in the array. The first array element becomes the last, and the last array element becomes the first.

## Syntax

*expression*.**Reverse**([*recurse*])

### Parameters

Name 
: `recurse`

Type
: `Boolean`

Necessity
: Optional

Description
: If true and the array is jagged or multi-dimensional (which are stored as jagged internally), the order of all nested arrays will also be reversed.

### Returns

Type
: `BetterArray` / `Object`

Description
: The current instance of the BetterArray object with the array's order reversed.


## Inspiration
* <http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.reverse>
* <https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/reverse>

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)