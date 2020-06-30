---
title: FromCSV
parent: Methods
grand_parent: API
---

# FromCSV

## Description

The `FromCSV()` method accepts a path argument pointing to a comma-separated values (CSV) file and stores the delimited values contained within to the internal array. Fields that contain a special character (comma, CR*, LF*, or double quote*), and are "escaped" by enclosing them in double quotes (Hex 22) are correctly handled as per the [RFC 4180](https://tools.ietf.org/html/rfc4180#page-2) specification. *WIP

    ByRef Path As String, _
    Optional ByVal ColumnDelimiter As String = ",", _
    Optional ByRef RowDelimiter As String = vbNewLine, _
    Optional ByRef Quote As String = """", _
    Optional ByVal IgnoreFirstRow As Boolean _


## Syntax

*expression*.**FromCSV**(*Path*, [*ColumnDelimiter*], [*RowDelimiter*], [*Quote*], [*IgnoreFirstRow*])

### Parameters

Name 
: `Path`

Type
: `String`

Necessity
: Required

Description
: A valid path string to the target CSV file.

---

Name 
: `ColumnDelimiter`

Type
: `String`

Necessity
: Optional

Description
: The character used to delimit columns within the CSV file. If omitted, the character `,` (comma) is used.

---

Name 
: `RowDelimiter`

Type
: `String`

Necessity
: Optional

Description
: The character(s) used to delimit rows within the CSV file. If omitted, the character stored in the constant `vbNewLine` is used.

---

Name 
: `Quote`

Type
: `String`

Necessity
: Optional

Description
: The character(s) used to escape characters within cells of the CSV file. If omitted, the character `"` (double quote) is used indicate the opening and closing of an escape sqeuence.

---

Name 
: `IgnoreFirstRow`

Type
: `Boolean`

Necessity
: Optional

Description
: If true, the first line of the CSV file will be skipped. Use this if your data has headers but you just want to return the data body.



### Returns

Type
: `BetterArray` / `Object`

Description
: The current instance of the BetterArray object with the data from the parsed CSV stored in the internal array.


## Example

```vb
Public Sub FromCSVExample()
    Dim MyArray As BetterArray
    Set MyArray = New BetterArray
    
    Dim path As String
    path = Strings.Join(Array(Environ("USERPROFILE"), "Desktop", "Data", "Sales Records.csv"), "\")
    
    Dim outputSheet As Worksheet
    Set outputSheet = ThisWorkbook.Sheets.Add
    MyArray.FromCSV(path).ToExcelRange outputSheet.Range("A1")
    
    ' expected output:
    ' The data in the CSV was parsed into an array and written to a new worksheet in Excel
End Sub
```


[Back to Docs](https://senipah.github.io/VBA-Better-Array/)





