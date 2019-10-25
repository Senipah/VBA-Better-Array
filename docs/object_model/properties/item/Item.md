---
title: Item
parent: Properties
has_children: true 
---

# Item

Gets or sets the element at the specified index.

## Note
`Item` is the default member of the `BetterArray` class. Subsequently the accessor can be invoked either explicitly:

```vb
MyArray.Item(10) = "Foo" ' Assigns the element at index 10 the value "Foo"
```

or Implicitly

```vb
MyArray(10) = "Foo" ' Assigns the element at index 10 the value "Foo"
```
