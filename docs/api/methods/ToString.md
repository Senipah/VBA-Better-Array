---
title: ToString
parent: Methods
grand_parent: API
---

# ToString

## Description
The `ToString()` method returns a string representing the array structure and its elements. Arrays are enclosed with square brackets. Elements are comma-separated. Object types are represented as "[Object]". Set the `PrettyPrint` argument to `True` to format the returned string with indentation for easier viewing of long or nested arrays.

## Syntax

*expression*.**ToString**([*prettyPrint*], [*delimitStrings*], [*openingDelimiter*], [*closingDelimiter*])

### Parameters

Name 
: `prettyPrint`

Type
: `Boolean`

Necessity
: Optional

Description
: Set to true to format the returned string for easier viewing.

---

Name 
: `delimitStrings`

Type
: `Boolean`

Necessity
: Optional

Description
: If true, any string values stored in the array will additionally be enclosed by opening and closing quotation marks (`"`).

---

Name 
: `openingDelimiter`

Type
: `String`

Necessity
: Optional

Description
: If provided, the string passed to `openingDelimiter` will be used to mark the beginning of arrays. If ommitted, the default character of `{` will be used.

---

Name 
: `closingDelimiter`

Type
: `String`

Necessity
: Optional

Description
: If provided, the string passed to `closingDelimiter` will be used to mark the beginning of arrays. If ommitted, the default character of `}` will be used.

### Returns

Type
: `String`

Description
: A string representing the array structure and its elements.

## Example

```vb
Public Sub ToStringExample()
    Dim result As String
    Dim MyArray As BetterArray
    Set MyArray = New BetterArray
    
    MyArray.Push "Banana", "Orange", "Apple", "Mango"
    result = MyArray.ToString
    Debug.Print result
    ' expected output:
    ' result = "{Banana,Orange,Apple,Mango}"
End Sub
```

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)