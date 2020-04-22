---
title: FromExcelRange
parent: Methods
grand_parent: API
---

# FromExcelRange

## Description

The `FromExcelRange()` method accepts an Excel Range object as an argument and stores the values contained within the range to the internal array. If the Range object contains only one row or column, the values will be stored and returned as a one-dimension array. 

If the Range contains both multiple Rows and Columns the default behaviour is that the array returned by accessing the `Items` get accessor will be returned as a multi-dimension array. 

If `detectLastColumn` is set to true, the range will be expanded along the first row of the `fromRange` until the last column containing data is found. If `detectLastRow` is set to true, the range will be expanded down the first column of the `fromRange` until the last row containing data is found. 

For example, assume you have a table of data starting at "A2" of a worksheet with 100 rows and 50 columns of data. Supplying the argument:

```vb
MyArray.FromExcelRange ActiveSheet.Range("A2"), True
```

Would store in `MyArray` a one-dimensional array with 100 entries storing the data in the range `"A2:A101"`, representing the full first column of data in the table.

Supplying the argument:

```vb
MyArray.FromExcelRange ActiveSheet.Range("A2"), False, True
```

Would store in `MyArray` a one-dimensional array with 50 entries storing the data in the range `"A2:AX2"`, representing the full first row of data in the table.

Supplying the argument:

```vb
MyArray.FromExcelRange ActiveSheet.Range("A2"), True, True
```

Would store in `MyArray` a two-dimensional array with all entries in the range `"A2:AX101"`, representing the full table.

## Syntax

*expression*.**FromExcelRange**(*fromRange*, [*detectLastRow*], [*detectLastColumn*])

### Parameters

Name 
: `fromRange`

Type
: `Range` / `Object`

Necessity
: Required

Description
: An Excel Range object containing the values to be stored in the array.

---

Name 
: `detectLastRow`

Type
: `Boolean`

Necessity
: Optional

Description
: If present, the range will be expanded along the first column in the range until the last used row containing data is found.

---

Name 
: `detectLastColumn`

Type
: `Boolean`

Necessity
: Optional

Description
: If present, the range will be expanded along the first row in the range until the last used column containing data is found.

### Returns

Type
: `BetterArray` / `Object`

Description
: The current instance of the BetterArray object with the values from the `fromRange` argument (expanded as determined by the `detectLastColumn` and `detectLastColumn` arguments) stored in the internal array.

## Example

```vb
Public Sub FromExcelRangeExample()
    Dim MyArray As BetterArray
    Set MyArray = New BetterArray
    
    MyArray.FromExcelRange ActiveSheet.UsedRange

    ' expected output:
    ' MyArray now stores the values in the UsedRange of the active worksheet
End Sub
```


[Back to Docs](https://senipah.github.io/VBA-Better-Array/)





