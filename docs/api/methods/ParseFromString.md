---
title: ParseFromString
parent: Methods
grand_parent: API
---

# ToString

## Description
The `ParseFromString()` method is used to deserialize a string representation of an array, such as that created by the `ToString()` method, back into an array.

        ByVal sourceString As String, _
        Optional ByVal valueSeparator As String = CHR_COMMA, _
        Optional ByVal arrayOpenDelimiter As String, _
        Optional ByVal arrayClosingDelimiter As String _

## Syntax

*expression*.**ParseFromString**(*sourceString*, [*valueSeparator*], [*arrayOpenDelimiter*], [*arrayClosingDelimiter*])

### Parameters

Name
: `sourceString`

Type
: `String`

Necessity
: Required

Description
: The string to be parsed.

---

Name
: `valueSeparator`

Type
: `String`

Necessity
: Optional

Description
: The character used to separate fields in the array. By default the comma (`,`) symbol is used.

---

Name
: `arrayOpenDelimiter`

Type
: `String`

Necessity
: Optional

Description
: If provided, the string passed to `openingDelimiter` will be used to mark the beginning of arrays. If ommitted, the default character of `{` will be used.

---

Name
: `arrayClosingDelimiter`

Type
: `String`

Necessity
: Optional

Description
: If provided, the string passed to `closingDelimiter` will be used to mark the beginning of arrays. If ommitted, the default character of `}` will be used.


### Returns

Type
: `BetterArray` / `Object`

Description
: The current instance of the BetterArray object with the values from the `sourceString` argument stored in the internal array.


## Example

```vb
Public Sub ParseFromString()
    Const ArrayString As String = "{Banana,Orange,Apple,Mango}"
    Dim myArray As BetterArray
    Set myArray = New BetterArray

    myArray.ParseFromString ArrayString
End Sub
```

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)
