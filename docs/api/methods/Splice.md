---
title: Splice
parent: Methods
grand_parent: API
---

# Splice

## Description
The `Splice()` method changes the contents of an array by removing or replacing existing elements and/or adding new elements in place.

## Syntax

*expression*.**Splice**(*StartIndex*, [*DeleteCount*,[*Item1*[, *Item2*[, ...[, *ItemN*]]]])

### Parameters

Name
: `StartIndex`

Type
: `Long`

Necessity
: Required

Description
: The index at which to start changing the array.
If greater than the length of the array, start will be set to the length of the array. In this case, no element will be deleted but the method will behave as an adding function, adding as many element as provided in the `Items` ParamArray argument.
A negative index can be used, indicating an offset from the end of the sequence

---

Name
: `DeleteCount`

Type
: `Long`

Necessity
: Required

Description
: An integral indicating the number of elements in the array to remove from start.
If DeleteCount is 0 or negative, no elements are removed. In this case, you should specify at least one new element (see below).
If DeleteCount is omitted, or if its value is equal to or larger the array's (`UpperBound` - `StartIndex`), then all the elements from start to the end of the array will be deleted.

---

Name
: `Items`

Type
: ParamArray `Variant`

Necessity
: Optional

Description
: The elements to add to the array, beginning from `StartIndex`. If you do not specify any elements, `Splice()` will only remove elements from the array.

### Returns

Type
: `Variant()`

Description
: A zero-based array containing the deleted elements. If only one element is removed, an array of one element is returned. If no elements are removed, an empty array is returned.

## Example

```vb
Public Sub SpliceExample()
    Dim result() As Variant
    Dim MyArray As BetterArray
    Set MyArray = New BetterArray

    MyArray.Push "Banana", "Orange", "Apple", "Mango"
    MyArray.Splice 2, 0, "Lemon", "Kiwi"
    result = MyArray.Items
    ' expected output:
    ' result =  "Banana", "Orange", "Lemon", "Kiwi", "Apple", "Mango"
End Sub

```

## Inspiration
* <http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.splice>
* <https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/Splice>


[Back to Docs](https://senipah.github.io/VBA-Better-Array/)
