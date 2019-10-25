---
title: Filter
parent: Methods
---


# Filter
### Inspiration
* http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.filter
* https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/filter
* https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/filter-function

### Description
The `Filter()` method filters and returns the current array LowerBound on the specified filter criteria. 

### Note
Provides ability to filter on variant arrays (not just with strings unlike the VBA version).

### Arguments
#### `Match` (Variant)
The value to compare against
#### *Optional* `Exclude` (Boolean)
Boolean value indicating whether to return values that include or exclude `Match`. If include is True, `Filter` returns the subset of the array that contains `Match`. If include is False, `Filter` returns the subset of the array that does not contain `Match`.
### Returns
#### (Variant)
The modified array.

# [Back to Docs](https://senipah.github.io/VBA-Better-Array/)