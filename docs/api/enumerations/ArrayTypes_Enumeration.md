---
title: ArrayTypes
parent: Enumerations
grand_parent: API
---

# ArrayTypes Enumeration

Specifies the structure of the stored array.

| Name              | Value | Description                                        |
|-------------------|-------|----------------------------------------------------|
| BA_UNDEFINED      | 0     | Array is Empty                                     |
| BA_UNALLOCATED    | 1     | Array has one slot containing a single Empty value |
| BA_ONEDIMENSION   | 2     | Valid array with only a single rank.               |
| BA_MULTIDIMENSION | 3     | Array is multidimensional (e.g. (0 To 1, 0 To 1))  |
| BA_JAGGED         | 4     | Array of Arrays                                    |
