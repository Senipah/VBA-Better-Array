---
title: Push
parent: Methods
grand_parent: API
---

# Push

## Description
The `Push()` method adds one or more elements to the end of an array, in the order in which they appear. The new length of the array is returned as the result of the call.

#### Note
Internal array capacity will be doubled automatically every time the length of the array equals the capacity of the internal array.

## Syntax

*expression*.**Push**([*args1*[, *args2*[, ...[, *argsN*]]]])

### Parameters

Name 
: `args`

Type
: ParamArray `Variant`

Necessity
: Optional

Description
: The element(s) to be added to the end of the array.

### Returns

Type
: `Long`

Description
: The new length of the array.

## Inspiration
* <https://developer.mozilla.org/en/docs/Web/JavaScript/Reference/Global_Objects/Array/push>
* <http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.push>


[Back to Docs](https://senipah.github.io/VBA-Better-Array/)