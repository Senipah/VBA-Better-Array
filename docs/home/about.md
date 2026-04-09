---
title: About
parent: Home
nav_order: 1
---

# About

The *BetterArray* class stores all arrays internally as `Variant()` dynamically allocated arrays. It supports one-dimension, multi-dimension and jagged arrays. If you assign the `Items` property a multi-dimension array, the class instance will first make a note that the array should be returned as a multi-dimension array and then convert it internally to a jagged array. This allows you to modify the shape, structure, order and contents of the array with ease using any of the built-in methods whilst still returning you a multi-dimension array.

The internal array inside the `BetterArray` instance has its own capacity, which is separate from the length of the array which will be returned to you. Similar to the processes in C#'s [ArrayList](https://docs.microsoft.com/en-us/dotnet/api/system.collections.arraylist?view=netframework-4.8) or GoLang's [append](https://golang.org/pkg/builtin/#append), each time the current internal capacity is reached it will be doubled. This should provide a small performance benefit over resizing the array each time a new element is added.

VBA Better Array includes a built-in VBA test harness and PowerShell runner scripts, so unit tests can be executed without external test add-ins.

VBA Better Array is free software; you can redistribute it and/or modify it under the terms of the MIT license for use in commercial or personal projects. See [LICENSE](https://github.com/Senipah/VBA-Better-Array/blob/master/LICENSE) for details.

