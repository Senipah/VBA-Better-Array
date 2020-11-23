---
title: ToCSVString
parent: Methods
grand_parent: API
---

# ToCSVString

## Description

The `ToCSVString()` returns an [RFC 4180](https://tools.ietf.org/html/rfc4180#section-2) compliant string representation of the stored array.

## Syntax

*expression*.**ToCSVString**([*Headers*], [*ColumnDelimiter*], [*RowDelimiter*], [*Quote*], [*EncloseAllInQuotes*], [*DateFormat*], [*NumberFormat*])

### Parameters

Name
: `Headers`

Type
: `Variant`

Necessity
: Optional

Description
: If provided, `Headers` should be a Variant() array containing header strings. The number of headers should be the same as the number of columns in the array.

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
: `EncloseAllInQuotes`

Type
: `Boolean`

Necessity
: Optional

Description
: If true, all fields will be enclosed by opening and closing characters. The string passed to the `Quote` argument is used (default is `"`).

---

Name
: `DateFormat`

Type
: `String`

Necessity
: Optional

Description
: If provided, all date values will be formmatted using the value in this argument forwarded to the VBA [Format()](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/format-function-visual-basic-for-applications) function.

---

Name
: `NumberFormat`

Type
: `String`

Necessity
: Optional

Description
: If provided, all numeric values will be formmatted using the value in this argument forwarded to the VBA [Format()](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/format-function-visual-basic-for-applications) function.




### Returns

Type
: `String`

Description
: The string CSV compatible string representation of the array.



[Back to Docs](https://senipah.github.io/VBA-Better-Array/)





