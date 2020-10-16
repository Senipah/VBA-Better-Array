---
title: FromCSVString
parent: Methods
grand_parent: API
---

# FromCSVString

## Description

The `FromCSVString()` method accepts a string argument representing the content of a comma-separated values (CSV) file and stores the delimited values contained within to the internal array. Fields that contain a special character (comma, CR\*, LF\*, or double quote\*), and are "escaped" by enclosing them in double quotes (Hex 22) are correctly handled as per the [RFC 4180](https://tools.ietf.org/html/rfc4180#page-2) specification.

Behaves the same as the `FromCSVString` method.

## Syntax

*expression*.**FromCSVString**(*CSVString*, [*ColumnDelimiter*], [*RowDelimiter*], [*Quote*], [*IgnoreFirstRow*], [*DuckType*])

### Parameters

Name 
: `CSVString`

Type
: `String`

Necessity
: Required

Description
: A string argument representing the content of a comma-separated values (CSV) file.

---

Name 
: `ColumnDelimiter`

Type
: `String`

Necessity
: Optional

Description
: The character used to delimit columns within the CSV file. If omitted, the character `,` (comma) is used.

---

Name 
: `RowDelimiter`

Type
: `String`

Necessity
: Optional

Description
: The character(s) used to delimit rows within the CSV file. If omitted, the character stored in the constant `vbNewLine` is used.

---

Name 
: `Quote`

Type
: `String`

Necessity
: Optional

Description
: The character(s) used to escape characters within cells of the CSV file. If omitted, the character `"` (double quote) is used indicate the opening and closing of an escape sqeuence.

---

Name 
: `IgnoreFirstRow`

Type
: `Boolean`

Necessity
: Optional

Description
: If true, the first line of the CSV file will be skipped. Use this if your data has headers but you just want to return the data body.

---

Name 
: `DuckType`

Type
: `Boolean`

Necessity
: Optional

Description
: If true, numeric and boolean values will be correctly converted to the appropriate type. If false all values will be String. Leave false if you just intend to output the values to an Excel worksheet as Excel will perform the type conversion automatically.




### Returns

Type
: `BetterArray` / `Object`

Description
: The current instance of the BetterArray object with the data from the parsed CSV stored in the internal array.



[Back to Docs](https://senipah.github.io/VBA-Better-Array/)





