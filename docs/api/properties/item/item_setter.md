---
title: Item Letter
parent: Item
grand_parent: Properties
nav_order: 2
---

# Item Let Accessor

## Description
Replaces the array item at the specified index with the passed element. If the `Index` argument exceeds the current upper bound of the stored array the element will be pushed onto the end of the array at the next available index (this changes the length and upper bound of the stored array). If the index argument is less than the LowerBound (base index) of the stored array the element will be inserted at the beginning of the array and the existing elements shifted up to accomodate the new element (this changes the length and upper bound of the stored array).

## Parameters
### `Index` (long) 
The index of the element in the array to replace.
### `Element` (variant) 
The element to be inserted into the array.

## Returns
**None**

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)