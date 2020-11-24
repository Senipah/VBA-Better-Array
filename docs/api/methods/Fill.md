---
title: Fill
parent: Methods
grand_parent: API
---


# Fill

## Description
The `Fill()` method method fills (modifies) all the elements of an array from a start index (default zero) to an end index (default array length) with a static value.

## Syntax

*expression*.**Fill**(*value*, [*StartIndex*], [*EndIndex*])

### Parameters

Name
: `value`

Type
: `Variant`

Necessity
: Required

Description
: The value to populate the array with.

---

Name
: `StartIndex`

Type
: `Long`

Necessity
: Optional

Description
: The first index of the outermost array to begin filling with the passed value. If ommited the array will be filled from the first index.

---

Name
: `EndIndex`

Type
: `Boolean`

Necessity
: Optional

Description
: The last index of the outermost array to to fill up to with the passed value.

### Returns

Type
: BetterArray `Object`

Description
: The current instance of the BetterArray object array filled with the passed value between the specified indices. If ommited the array will be filled to the last index.

## Example

```vb
Public Sub FillExample()
    Dim MyArray As BetterArray
    Dim result() As Variant

    Set MyArray = New BetterArray
    MyArray.Push 1, 2, 3, 4
    MyArray.Fill 0, 2, 4
    result = MyArray.Items

    ' expected output:
    ' result is a array with the values: 1, 2, 0, 0
End Sub
```

## Inspiration
* <http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.fill>
* <https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/fill>

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)
