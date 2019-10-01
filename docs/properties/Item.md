# Item

## Get

### Description
Returns the element stored in the array at the specified index.

### Arguments
#### `index` (Long) 
The index of the desired element.

### Returns
#### (Variant) 
The element at the specified index.

## Let

### Description
Replaces the array item at the specified index with the passed element. If the `index` argument exceeds the current upper bound of the stored array the element will be pushed onto the end of the array at the next available index (this changes the length and upper bound of the stored array). If the index argument is less than the base (lower bound) of the stored array the element will be inserted at the beginning of the array and the existing elements shifted up to accomodate the new element (this changes the length and upper bound of the stored array).

### Arguments
#### `index` (long) 
The index of the element in the array to replace.
#### `element` (variant) 
The element to be inserted into the array.

### Returns
None

# [Back to Docs](https://senipah.github.io/VBA-DynamicArray/)