---
title: Getting Started
parent: Home
nav_order: 3
---

# Getting Started

If you don't already have the `BetterArray.cls` code added to your project, please refer to the [installation instructions](https://senipah.github.io/VBA-Better-Array/home/installation.html).

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

## Accessing elements in the array

To access individual elements stored in the array, use the [Item](https://senipah.github.io/VBA-Better-Array/api/properties/item/Item.html) property. As Item is the default member of the better array class, it can be accessed explicilty or implicitly.

### Retrieving elements

Explicit access:

```vb
Public Sub ElementAccess()
    Dim Foods As BetterArray
    Set Foods = New BetterArray
    Foods.Push "Cheese", "Eggs", "Ham"
    Debug.Print Foods.Item(1)
    ' expected output: "Eggs"
End Sub
```

Implicit access:

```vb
Public Sub ElementDefaultMemberAccess()
    Dim Foods As BetterArray
    Set Foods = New BetterArray
    Foods.Push "Cheese", "Eggs", "Ham"
    Debug.Print Foods(1)
    ' expected output: "Eggs"
End Sub
```

### Updating elements

Similarly, the Items property is used to change the element stored in the array at a specified index.

Explicit access:

```vb
Public Sub ElementAccess()
    Dim Foods As BetterArray
    Set Foods = New BetterArray
    Foods.Push "Cheese", "Eggs", "Ham"
    Foods.Item(1) = "Steak"
    ' new array values: "Cheese", "Steak", "Ham"
End Sub
```

Implicit access:

```vb
Public Sub ElementDefaultMemberAccess()
    Dim Foods As BetterArray
    Set Foods = New BetterArray
    Foods.Push "Cheese", "Eggs", "Ham"
    Foods(1) = "Steak"
    ' new array values: "Cheese", "Steak", "Ham"
End Sub
```

#### Note

If you try to assign an element to an index which exceeds the current [UpperBound](https://senipah.github.io/VBA-Better-Array/api/properties/upper_bound/UpperBound.html) of the array, the element will be assigned at the next available index in the array.

Example:

```vb
Public Sub ElementDefaultMemberAccess()
    Dim Foods As BetterArray
    Set Foods = New BetterArray
    Foods.Push "Cheese", "Eggs", "Ham"
    Foods(10) = "Steak"
    ' new array values: "Cheese", "Eggs", "Ham", "Steak"
    ' new array bounds: 0 to 3
End Sub
```

## Assigning a whole array

You can import an array into the BetterArray instance by assiging it with the [Items](https://senipah.github.io/VBA-Better-Array/api/properties/items/Items.html) property let accessor.

Example:

```vb
Public Sub AssigningAnArray()
    Dim originalArray(0 To 2) As Variant
    Dim Foods As BetterArray
    Set Foods = New BetterArray

    originalArray(0) = "Cheese"
    originalArray(1) = "Eggs"
    originalArray(2) = "Ham"
        
    Foods.Items = originalArray
    Foods.Push "Steak"
    ' Foods is now the following list: "Eggs", "Cheese", "Ham", "Steak"
End Sub
```

## Retrieving the stored array

Accessing the [Items](https://senipah.github.io/VBA-Better-Array/api/properties/items/Items.html) property get accessor will return the stored array to you as a `Variant` array.

```vb
Public Sub RetrievingAnArray()
    Dim myShoppingList() As Variant
    Dim Foods As BetterArray
    Set Foods = New BetterArray

    Foods.Push "Cheese"
    Foods.Push "Eggs"
    Foods.Push "Ham"
        
    myShoppingList = Foods.Items
End Sub
```
## Iterating over the array

The Items property returns an array which is inherently iterable and can be used in a for each if you just want to retrieve all of the values. To mutate all of the values in a better array, iterate by index: 

```vb
Public Sub Iterating_PlusOne()
    Dim Numbers As BetterArray
    Set Numbers = New BetterArray
    Numbers.Push 1, 2, 3
    Dim i As Long
    For i = Numbers.LowerBound To Numbers.UpperBound
        Numbers(i) = Numbers(i) + 1
    Next
    ' Numbers is now: 2, 3, 4
End Sub
```
