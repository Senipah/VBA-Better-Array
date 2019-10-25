---
title: LowerBound
parent: Properties
grand_parent: API
---

# LowerBound

The `LowerBound` property stores the starting index of the array. By default, the lower bound will be set to 0. 

### Note 

If the `LowerBound` property has not been set by a user and the internal array is assigned using the `Items` let accessor, the `LowerBound` value of the assigned array will be used. If it has been user-specified before an array is assigned with the `Items` let accessor, the passed array will be re-indexed to start at the user-specified starting index.

If the `BetterArray` instance already contains elements in its internal array the internal array will be re-index to begin at the new `LowerBound` value.


[Back to Docs](https://senipah.github.io/VBA-Better-Array/)
