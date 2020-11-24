---
title: CopyFromCollection
parent: Methods
grand_parent: API
---

# CopyFromCollection

## Description
The `CopyFromCollection()` method writes the elements stored in a `Collection` object into a BetterArray. Any items already stored in the BetterArray will be lost.

## Syntax

*expression*.**Concat**(*SourceCollection*)

### Parameters

Name
: `SourceCollection`

Type
: `Collection` / `Object`

Necessity
: Required

Description
: The `Collection` of elements to be stored in the array

### Returns

Type
: `BetterArray` / `Object`

Description
: The current instance of the BetterArray object with the values from the passed `Collection` stored in the array.

## Example

```vb
Public Sub CopyFromCollectionExample()
    Dim MyCollection As Collection
    Dim MyArray As BetterArray
    Dim result() As Variant

    Set MyCollection = New Collection
    Set MyArray = New BetterArray

    MyCollection.Add "Foo"
    MyCollection.Add "Bar"
    MyCollection.Add "Fizz"
    MyCollection.Add "Buzz"

    MyArray.CopyFromCollection MyCollection

    result = MyArray.Items

    ' expected output:
    ' result is an array with the values: "Foo", "Bar", "Fizz", "Buzz"
End Sub
```



[Back to Docs](https://senipah.github.io/VBA-Better-Array/)
