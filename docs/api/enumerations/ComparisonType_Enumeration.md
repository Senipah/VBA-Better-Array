---
title: ComparisonType
parent: Enumerations
grand_parent: API
---

# ComparisonType Enumeration

Specifies the type of comparison to be performed.

| Name        | Value | Description                                                                                                                                                                                                                |
|-------------|-------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| CT_EQUALITY | 0     | Compares `searchElement` against the element at the current index for equality. This is the default comparison method.                                                                                                     |
| CT_LIKENESS | 1     | `searchElement` is treated as a string pattern and compared against the element as the current index using the `Like` operator. If this option is chosen `searchElement` must be a String type or an error will be raised. |
