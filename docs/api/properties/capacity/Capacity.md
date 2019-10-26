---
title: Capacity
parent: Properties
grand_parent: API
has_children: true 
---

# Capacity

The `Capacity` property is used to used to get or set the number of elements that the internal array can contain. If the size of the array is known ahead of time then setting the capacity may result in a small performance boost when assigning elements; otherwise, the capacity will be set as required when new elements are added.

#### Note
Similar to the processes in C#'s [ArrayList](https://docs.microsoft.com/en-us/dotnet/api/system.collections.arraylist?view=netframework-4.8) or GoLang's [append](https://golang.org/pkg/builtin/#append), each time the current internal capacity is reached it will be doubled. This should provide a small performance benefit over resizing the array each time a new element is added.

