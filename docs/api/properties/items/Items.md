---
title: Items
parent: Properties
grand_parent: API
has_children: true 
---

# Items

Assigns or retrieves the stored array. Assigned arrays must be of type `Variant`.

#### Note
Multi-dimensional arrays are converted internally to jagged arrays. If the array was initiated by passing a multi-dimensional array to the Let property, the array returned by the Get accessor will be converted from a jagged array to a multi-dimensional array.

If a non-array type is passed to Items then the existing array will be cleared and the argument will be pushed to the now empty array as the first element.
