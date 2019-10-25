---
title: Max
parent: Methods
grand_parent: API
---

# Max

## Description
The `Max()` method returns the largest value in a list of values.  If no argments are passed the `Max()` method will use the stored array. Returns `Empty` if array is uninitialized, or only contains non-scalar variables. If operating on a jagged or multi-dimensional array, the value returned will be the largest value in all of the arrays combined.
  
### Note 
Multi-dimensional arrays assigned to the `.Items` property are converted to jagged arrays internally and will be treated as such by the `Max()` method.

## Syntax

*expression*.**Max**([*Args1*[, *Args2*[, ...[, *ArgsN*]]]])

### Parameters

Name 
: `Args`

Type
: ParamArray `Variant`

Necessity
: Opional

Description
: A list of values or an array to compare. If no arguments are provided the `Max()` method will return the largest value on the stored array.

### Returns

Type
: `Variant`

Description
: The largest value in the array.

## Inspiration
* <https://support.office.com/en-gb/article/max-function-e0012414-9ac8-4b34-9a47-73e662c08098>
* <https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Math/max>
* <http://www.ecma-international.org/ecma-262/10.0/index.html#sec-math.max>


[Back to Docs](https://senipah.github.io/VBA-Better-Array/)