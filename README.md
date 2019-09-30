# UNDER CONSTRUCTION


# Project name: VBA-DynamicArray

## Description: A VBA class providing a more modern and user-friendly Array.

## Table of Contents: Optionally, include a table of contents in order to allow other people to quickly navigate especially long or detailed READMEs.

## Properties

### Capacity

#### Get

##### Description
Returns the current capacity of the internal array.
##### Arguments
None
##### Returns
###### (Long)
Current capacity of the internal array.

#### Let
##### Description
Sets the capacity of the internal array.
##### Arguments
###### `value` (Long) 
The new size of the internal array.
##### Returns
None

### Length

*(Read-Only)*

#### Get

##### Description
Gets the length of the current array
##### Arguments
None
##### Returns
###### (Long) 
The length of the current array.

### Items

#### Get
##### Description
Returns the current item list as a variant array
##### Arguments
None
##### Returns
###### (Variant) 
Variant array of the current items,

#### Let
##### Description
Stores an array into the DynamicArray object
##### Note 
*multi-dimensional arrays are converted internally to jagged arrays. If the array was initiated by passing a multi-dimensional array to the Let property, the array returned by the Get property will be converted from a jagged array to a multi-dimensional array.*
##### Arguments
###### `values` (Variant) 
The array of items to import to the DynamicArray object.
##### Returns
None

### UpperBound

*(Read-Only)*

#### Get
##### Description
Returns the upper bound of the array
##### Arguments
None
##### Returns
##### (Long) 
Upper bound of the array.

### Base

*(Read-Only)*

#### Get

##### Description
Gets the base (starting index) of the array

##### Arguments
None

##### Returns
###### (Long) 
Starting index of the array.

### Item

#### Get

##### Description
Returns the element stored in the array at the specified index.

##### Arguments
##### `index` (Long) 
The index of the desired element.

##### Returns
###### (Variant) 
The element at the specified index.

#### Let

##### Description
Replaces the array item at the specified index with the passed element. If the `index` argument exceeds the current upper bound of the stored array the element will be pushed onto the end of the array at the next available index (this changes the length and upper bound of the stored array). If the index argument is less than the base (lower bound) of the stored array the element will be inserted at the beginning of the array and the existing elements shifted up to accomodate the new element (this changes the length and upper bound of the stored array).

##### Arguments
###### `index` (long) 
The index of the element in the array to replace.
###### `element` (variant) 
The element to be inserted into the array.

##### Returns
None

## Methods

### Push
* https://developer.mozilla.org/en/docs/Web/JavaScript/Reference/Global_Objects/Array/push
* http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.push

##### Description
The `Push()` method adds one or more elements to the end of an array, in the order in which they appear. The new length of the array is returned as the result of the call.

##### Note
*Internal array capacity will be doubled automatically every time the length of the array equals or exceeds the capacity of the internal array.*

##### Arguments
###### ParamArray `args` (variant) 
The element(s) to be added to the end of the array.

##### Returns
###### (Long) 
New length of the array.

### Pop

##### Inspiration
* https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/pop
* http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.pop

##### Description
The `Pop()` method removes the last element from the array and returns that element. This method changes the length of the array.

##### Arguments
None

##### Returns
###### (Variant) 
The last element from the array.

### Shift

##### Inspiration
* https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/shift
* http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.shift

##### Description
The `Shift()` method removes the first element from the array and returns that removed element. This method changes the length of the array.

##### Arguments
None

##### Returns
###### (Variant) 
The first element from the array.

### Unshift

##### Inspiration
* https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/Unshift
* http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.unshift

##### Description
The `Unshift()` method adds one or more elements to the beginning of an array and returns the new length of the array.

##### Arguments
###### ParamArray `args`() (Variant) 
The elements to be added to the beginning of the array

##### Returns
###### (Long) 
New length of the array.

### Concat

##### Inspiration
* https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/concat
* http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.concat

##### Description
The `Concat()` method joins one or more arrays onto the end of the current array.

##### Arguments
###### ParamArray `args`() (Variant) 
The array(s) to be added to the end of the array

##### Returns
###### (Variant) 
The new array

### CopyFromCollection

##### Description
The `CopyFromCollection()` method writes the elements stored in a collection object into a DynamicArray. Any items already stored in the DynamicArray will be lost.

##### Arguments
###### `c` (Collection [Object]) 
The collection of elements to be stored in the array

##### Returns:
###### (Variant) 
The new array

### ToString

##### Description
The `ToString()` method returns a string representing the array structure and its elements. Arrays are enclosed with square brackets. Elements are comma-separated. Object types are represented as "[Object]". Set the `prettyPrint` argument to `True` to format the returned string with indentation for easier viewing of long or nested arrays.

##### Arguments
###### *Optional* `prettyPrint` (Boolean)
Set to true to format the returned string for easier viewing.

##### Returns
###### (String) 
A string representing the array structure and its elements.

### Sort

##### Description
The `Sort()` method sorts and returns the stored array. If the array has more than one dimension the `col` argument is used to determine the column in the array to be used for the comparison. Arrays with more than two dimensions are unsupported and will be returned unchanged. Uses an implementation of the [Quicksort](https://en.wikipedia.org/wiki/Quicksort) algorithm

##### Arguments
###### *Optional* `col` (Long)
The column in a two dimensional array to be used in the comparison.

##### Returns
##### (Variant)
The sorted array.

### CopyWithin
##### Inspiration
* http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.copywithin
* https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/copyWithin

##### Description
The `copyWithin()` method shallow copies part of an array to another location in the same array and returns it without modifying its length. The copyWithin method takes up to three arguments target, start and end.

##### Note
*The end argument is optional with the length of the this object as its default value. If target is negative, it is treated as length + target where length is the length of the array. If start is negative, it is treated as length + start. If end is negative, it is treated as length + end.*

##### Arguments
###### `target` (Long)
The index at which to copy the sequence to. If negative, `target` will be counted from the end.
If `target` is at or greater than the array's `Length` property, nothing will be copied. If `target` is positioned after `startI`, the copied sequence will be trimmed to fit the array's `Length` property.
###### *Optional* `startI` (Long)
The index at which to start copying elements from. If negative, `startI` will be counted from the end.
If `startI` is omitted, `copyWithin` will copy from the base index of the array. 
###### *Optional* `endI` (Long)
The index at which to end copying elements from. `copyWithin` copies up to but not including `endI`. If negative, `endI` will be counted from the end.
If `endI` is omitted, `copyWithin` will copy until the last index (default to the array's `Length` property).
##### Returns
##### (Variant)
The modified array.



### Filter
##### Inspiration
* http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.filter
* https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/filter
* https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/filter-function

##### Description
The `Filter()` method filters and returns the current array based on the specified filter criteria. 

##### Note
Provides ability to filter on variant arrays (not just with strings unlike the VBA version).

##### Arguments
###### `match` (Variant)
The value to compare against
###### *Optional* `exclude` (Boolean)
Boolean value indicating whether to return values that include or exclude `match`. If include is True, `Filter` returns the subset of the array that contains `match`. If include is False, `Filter` returns the subset of the array that does not contain `match`.
##### Returns
##### (Variant)
The modified array.


### Includes
##### Inspiration
* http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.includes
* https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/includes

##### Description
The `Includes()` method determines whether the array includes a certain value among its entries, returning `True` or `False` as appropriate.

##### Arguments
###### `searchElement` (Variant)
The value to search for.
###### *Optional* `fromIndex` (Long)
The position in this array at which to begin searching for `searchElement`; the first character to be searched is found at `fromIndex` for positive values of `fromIndex`, or at the array's `Length` property + `fromIndex` for negative values of `fromIndex` (using the absolute value of `fromIndex` as the number of characters from the end of the array at which to start the search). Defaults to the array's `Base` property.
##### Returns
##### (Boolean)
`True` if the array includes valueToFind, `False` if not.

### Keys
##### Inspiration
* http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.keys
* https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/keys

##### Description
The `Keys()` method returns a new array that contains the keys for each index in the array.

##### Arguments
None
##### Returns
##### (Variant)
A new array containing the keys of all the elements in the array. 

### Max
##### Inspiration
* https://support.office.com/en-gb/article/max-function-e0012414-9ac8-4b34-9a47-73e662c08098
* https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Math/max
* http://www.ecma-international.org/ecma-262/10.0/index.html#sec-math.max

##### Description
The `Max()` method returns the largest value in a list of values. If passed a nested array `Max()` will return the largest value in the first array.  If no argments are passed the `Max()` method will use the stored array. Returns `Empty` if array is uninitialized, or only contains non-scalar variables or is a multidimensional array (see Note).
  
##### Note 
Multidimensional arrays assigned to the `.Items` property are converted to jagged arrays internally and will be treated as such by the `Max()` method).

##### Arguments
###### ParamArray `args()` (Variant)
A list of values or an array to compare. If no arguments are provided the `Max()` method will return the largest value on the stored array.


##### Returns
##### (Variant)
The largest value in the array. 

### Min
##### Inspiration
* https://support.office.com/en-gb/article/min-function-61635d12-920f-4ce2-a70f-96f202dcc152

##### Description
The `Min()` method returns the smallest value in a list of values. If passed a nested array `Min()` will return the smallest value in the first array.  If no argments are passed the `Min()` method will use the stored array. Returns `Empty` if array is uninitialized, or only contains non-scalar variables or is a multidimensional array (see Note).

##### Note 
Multidimensional arrays assigned to the `.Items` property are converted to jagged arrays internally and will be treated as such by the `Min()` method.

##### Arguments
A list of values or an array to compare. If no arguments are provided the `Min()` method will return the smallest value on the stored array.
##### Returns
##### (Variant)
The largest value in the array. 


### Slice
##### Inspiration
* http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.slice
* https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/slice


##### Description
The `Slice()` method returns a shallow copy of a portion of an array into a new array object selected from begin to end (end not included) where begin and end represent the index of items in that array. The original array will not be modified.

##### Arguments
###### `startI` (Long)
The index at which to begin extraction.
A negative index can be used, indicating an offset from the end of the sequence. slice(-2) extracts the last two elements in the sequence.
If begin is undefined, slice begins from index 0.
If begin is greater than the length of the sequence, an empty array is returned.
###### *Optional* `endI` (Long)
The index before which to end extraction. slice extracts up to but not including end.
For example, slice(1,4) extracts the second element through the fourth element (elements indexed 1, 2, and 3).
A negative index can be used, indicating an offset from the end of the sequence. slice(2,-1) extracts the third element through the second-to-last element in the sequence.
If end is omitted, slice extracts through the end of the sequence (arr.length).
If end is greater than the length of the sequence, slice extracts through to the end of the sequence (arr.length).
##### Returns
##### (Variant)
A new array containing the extracted elements.

### FromRange
##### Inspiration
##### Description
##### Arguments
###### `Range` (Range [Object])
##### Returns
##### (Variant)






## Installation: Simply download DynamicArray.cls from this repo and import into your VBA project.

## Usage: The next section is usage, in which you instruct other people on how to use your project after they’ve installed it. This would also be a good place to include screenshots of your project in action.

## Contributing: Accepting pull requests.

## Credits: Include a section for credits in order to highlight and link to the authors of your project.

## License: This library is free software; you can redistribute it and/or modify it under the terms of the MIT license. See LICENSE for details.
