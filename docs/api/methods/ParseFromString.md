---
title: ParseFromString
parent: Methods
grand_parent: API
---

# ToString

## Description
The `ParseFromString()` method is used to deserialize a string representation of an array, such as that created by the `ToString()` method, back into an array.

## Syntax

*expression*.**ParseFromString**(*SourceString*, [*ValueSeparator*], [*ArrayOpenDelimiter*], [*ArrayClosingDelimiter*])

### Parameters

Name
: `SourceString`

Type
: `String`

Necessity
: Required

Description
: The string to be parsed.

---

Name
: `ValueSeparator`

Type
: `String`

Necessity
: Optional

Description
: The character used to separate fields in the array. By default the comma (`,`) symbol is used.

---

Name
: `ArrayOpenDelimiter`

Type
: `String`

Necessity
: Optional

Description
: If provided, the string passed to `ArrayOpenDelimiter` will be used to mark the beginning of arrays. If omitted, the default character of `{` will be used.

---

Name
: `ArrayClosingDelimiter`

Type
: `String`

Necessity
: Optional

Description
: If provided, the string passed to `ArrayClosingDelimiter` will be used to mark the beginning of arrays. If omitted, the default character of `}` will be used.


### Returns

Type
: `BetterArray` / `Object`

Description
: The current instance of the BetterArray object with the values from the `SourceString` argument stored in the internal array.


## Example

```vb
Public Sub ParseFromString()
    Const ArrayString As String = "{Banana,Orange,Apple,Mango}"
    Dim MyArray As BetterArray
    Set MyArray = New BetterArray

    MyArray.ParseFromString ArrayString
End Sub
```

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)
