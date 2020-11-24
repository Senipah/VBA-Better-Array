---
title: ToExcelRange
parent: Methods
grand_parent: API
---

# ToExcelRange

## Description

The `ToExcelRange()` method writes the values stored in the array to an excel worksheet starting at the specified range."


## Syntax

*expression*.**ToExcelRange**(*Destination*, [*TransposeValues*])

### Parameters

Name
: `Destination`

Type
: `Range` / `Object`

Necessity
: Required

Description
: An Excel Range object representing the loction to begin writing the stored values. Will be expanded as necessar to accomodate the size of the array.

---

Name
: `TransposeValues`

Type
: `Boolean`

Necessity
: Optional

Description
: If present, stored values will be transposed when written to the Excel Range (rows become columns and vice versa)

### Returns

Type
: `Range` / `Object`

Description
: The Excel Range object containing the outputted values.

## Example

```vb
Public Sub ToExcelRangeExample()
    Dim result() As Variant
    Dim MyArray As BetterArray
    Set MyArray = New BetterArray

    MyArray.Push "Banana", "Orange", "Apple", "Mango"
    MyArray.ToExcelRange ThisWorkbook.Sheets.Add.Range("A1"), True
    ' expected output:
    ' A new worksheet has been added and the MyArray items have been
    ' written to A1:A4
End Sub
```

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)





