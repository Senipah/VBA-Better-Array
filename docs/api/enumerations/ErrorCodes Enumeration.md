---
title: ErrorCodes
parent: Enumerations
grand_parent: API
---

# ErrorCodes Enumeration

Specifies the type and value of error.

| Name                                        | Value       | Description                                                                                 |
|---------------------------------------------|-------------|---------------------------------------------------------------------------------------------|
| EC_EXPECTED_RANGE_OBJECT                    | -2147220991 | Range Object Expected                                                                       |
| EC_EXPECTED_COLLECTION_OBJECT               | -2147220990 | Valid Collection Object Expected                                                            |
| EC_MAX_DIMENSIONS_LIMIT                     | -2147220989 | Cannot convert structure of arrays with more than 20 dimensions.                            |
| EC_EXCEEDS_MAX_SORT_DEPTH                   | -2147220988 | Cannot sort on arrays with more than 2 dimensions                                           |
| EC_EXPECTED_JAGGED_ARRAY                    | -2147220987 | Expected jagged array.                                                                      |
| EC_EXPECTED_MULTIDIMENSION_ARRAY            | -2147220986 | Expected multidimension array.                                                              |
| EC_EXPECTED_ARRAY                           | -2147220985 | Expected array.                                                                             |
| EC_NULL_STRING                              | -2147220984 | Cannot parse from a null string. Expected string with length greater than 0.                |
| EC_UNALLOCATED_ARRAY                        | -2147220983 | Cannot operate on unallocated array.                                                        |
| EC_UNDEFINED_ARRAY                          | -2147220982 | Array is undefined.                                                                         |
| EC_INVALID_MULTIDIMENSIONAL_ARRAY_OPERATION | -2147220981 | Unable to perform the requested operation on a multidimensional array.                      |
| EC_EXPECTED_VARIANT_ARRAY                   | -2147220980 | Unable to perform the requested operation on a typed array.                                 |
| EC_EXCEEDS_MAX_ARRAY_LENGTH                 | -2147220979 | The requested operation would result in an array which exceeds the maximum possible length. |
| EC_STRING_TYPE_EXPECTED                     | -2147220978 | Expected a String or String-coercible type.                                                 |
| EC_CANNOT_CONVERT_TO_REQUESTED_STRUCTURE    | -2147220977 | The stored array cannot be converted to the requested structure.                            |