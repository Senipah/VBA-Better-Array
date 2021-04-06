---
title: ArrayType Letter
parent: ArrayType
grand_parent: Properties
nav_order: 2
---

# ArrayType Let Accessor

## Description

Used to set the structure of the stored array to one of the valid [ArrayTypes Enumeration](https://senipah.github.io/VBA-Better-Array/api/enumerations/ArrayTypes%20Enumeration.html) options.

For example, if a multidimensional array is assigned to the BetterArray instance using the [Items setter](https://senipah.github.io/VBA-Better-Array/api/properties/items/items_setter.html), this can be converted into a jagged array (an array of arrays) by setting the `ArrayType` to `BA_JAGGED` (4) and then retrieving the stored array with the [Items getter](https://senipah.github.io/VBA-Better-Array/api/properties/items/items_getter.html). Similarly, a stored jagged array can be converted into a multidimensional array by setting `ArrayType` to `BA_MULTIDIMENSION` (3) prior to retrieval.

Attempting to set a value of `BA_UNDEFINED` will always throw a [EC_CANNOT_CONVERT_TO_REQUESTED_STRUCTURE](https://senipah.github.io/VBA-Better-Array/api/enumerations/ErrorCodes%20Enumeration.html) error.

Attempting to set a value of `BA_ONEDIMENSION` when the stored array is already either a multidimension or jagged array will throw a [EC_CANNOT_CONVERT_TO_REQUESTED_STRUCTURE](https://senipah.github.io/VBA-Better-Array/api/enumerations/ArrayTypes%20Enumeration.html) error.

Setting as value of `BA_UNALLOCATED` will clear the existing array. This is the same as calling the [ResetToDefault](https://senipah.github.io/VBA-Better-Array/api/methods/ResetToDefault.html) method.


## Syntax

*expression*.**ArrayType** = *NewType*

### Parameters

Name
: `NewType`

Type
: `ArrayTypes`/`Long`

Necessity
: Required

Description
: The desired [ArrayTypes Enumeration](https://senipah.github.io/VBA-Better-Array/api/enumerations/ArrayTypes%20Enumeration.html) value.


### Returns

**None**

[Back to ArrayType overview](https://senipah.github.io/VBA-Better-Array/api/properties/array_type/ArrayType.html)
