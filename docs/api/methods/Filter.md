---
title: Filter
parent: Methods
grand_parent: API
---


# Filter

## Description
The `Filter()` method filters and returns the current array based on the specified filter criteria.

#### Note
Provides ability to filter on any type of array (unlike the VBA version which only works with `String` arrays).

## Syntax

*expression*.**Filter**(*Match*, [*Exclude*], [*Recurse*])

### Parameters

Name
: `Match`

Type
: `Variant`

Necessity
: Required

Description
: The value to compare against.

---

Name
: `Exclude`

Type
: `Boolean`

Necessity
: Optional

Description
: Boolean value indicating whether to return values that include or exclude `match`. If include is True, `Filter` returns the subset of the array that contains `match`. If include is False, `Filter` returns the subset of the array that does not contain `match`.

---

Name
: `Recurse`

Type
: `Boolean`

Necessity
: Optional

Description
: Boolean value indicating whether the filter should be applied recrsively to a jagged or multidimensional array. By default, only the outermost array will be filtered.

### Returns

Type
: BetterArray `Object`

Description
: The current instance of the BetterArray object with the filter applied to the stored array.

## Example

```vb
Public Sub FilterExample()
    Dim MyArray As BetterArray
    Dim result() As Variant
    Set MyArray = New BetterArray

    MyArray.Push "Foo", "Bar", "Fizz", "Buzz"
    MyArray.Filter "Bar"

    result = MyArray.Items

    ' expected output:
    ' result is a array with the values: "Foo", "Fizz", "Buzz"
End Sub
```

## Inspiration
* <http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.filter>
* <https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/filter>
* <https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/filter-function>


[Back to Docs](https://senipah.github.io/VBA-Better-Array/)
