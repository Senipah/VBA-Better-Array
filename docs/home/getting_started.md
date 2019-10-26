---
title: Getting Started
parent: Home
nav_order: 3
---

# Getting Started

If you don't already have the `BetterArray.cls` code added to your project, please refer to the [Installation instructions](https://senipah.github.io/VBA-Better-Array/home/installation.html).

## Creating your first BetterArray instance

A better array instance is created like any other object variable in VBA.

```vb
Dim MyArray as BetterArray
Set MyArray = New BetterArray
```

`MyArray` is now an new instance of the `BetterArray` Class.

## Adding items to the array.

The simplest way to add items to `BetterArray` is to use the [Push](https://senipah.github.io/VBA-Better-Array/api/methods/Push.html) method.

Items can be added either one at a time:

```vb
Public Sub PushIndividual()
    Dim Foods As BetterArray
    Set Foods = New BetterArray
    
    Foods.Push "Cheese"
    Foods.Push "Eggs"
    Foods.Push "Ham"
End Sub
```

Or you can add multiple entries at the same time:

```vb
Public Sub PushMultiple()
    Dim Foods As BetterArray
    Set Foods = New BetterArray
    
    Foods.Push "Cheese", "Eggs", "Ham"
End Sub
```
