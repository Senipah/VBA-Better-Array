---
title: Min
parent: Methods
grand_parent: API
---


# Min

## Description
The `Min()` method returns the smallest value in a list of values.  If no argments are passed the `Min()` method will use the stored array. Returns `Empty` if array is uninitialized, or only contains non-scalar variables. If operating on a jagged or multi-dimensional array, the value returned will be the smallest value in all of the arrays combined.
  
#### Note
Multi-dimensional arrays assigned to the `.Items` property are converted to jagged arrays internally and will be treated as such by the `Min()` method.

## Syntax

*expression*.**Min**([*args1*[, *args2*[, ...[, *argsN*]]]])

### Parameters

Name 
: `args`

Type
: ParamArray `Variant`

Necessity
: Opional

Description
: A list of values or an array to compare. If no arguments are provided the `Min()` method will return the smallest value on the stored array.

### Returns

Type
: `Variant`

Description
: The smallest value in the array.

* <https://support.office.com/en-gb/article/min-function-61635d12-920f-4ce2-a70f-96f202dcc152>
* <https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Math/min>
* <http://www.ecma-international.org/ecma-262/10.0/index.html#sec-math.min>


[Back to Docs](https://senipah.github.io/VBA-Better-Array/)

