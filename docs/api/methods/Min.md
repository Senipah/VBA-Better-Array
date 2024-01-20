---
title: Min
parent: Methods
grand_parent: API
---


# Min

## Description
The `Min()` method returns the smallest value in a list of values.  If no arguments are passed the `Min()` method will use the stored array. Returns `Empty` if array is uninitialized, or only contains non-scalar variables. If operating on a jagged or multi-dimensional array, the value returned will be the smallest value in all of the arrays combined.

#### Note
Multi-dimensional arrays assigned to the `.Items` property are converted to jagged arrays internally and will be treated as such by the `Min()` method.

## Syntax

*expression*.**Min**([*Args1*[, *Args2*[, ...[, *ArgsN*]]]])

### Parameters

Name
: `Args`

Type
: ParamArray `Variant`

Necessity
: Optional

Description
: A list of values or an array to compare. If no arguments are provided the `Min()` method will return the smallest value on the stored array.

### Returns

Type
: `Variant`

Description
: The smallest value in the array.

## Example

```vb
Public Sub MinExample()
    Dim result As Long
    Dim MyArray As BetterArray
    Set MyArray = New BetterArray

    MyArray.Push 10, 1, 3, 5, 9, 12, 2, 8, 7
    result = MyArray.Min
    ' expected output:
    ' result  = 1
End Sub
```

## Inspiration
* <https://support.office.com/en-gb/article/min-function-61635d12-920f-4ce2-a70f-96f202dcc152>
* <https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Math/min>
* <http://www.ecma-international.org/ecma-262/10.0/index.html#sec-math.min>


[Back to Docs](https://senipah.github.io/VBA-Better-Array/)
