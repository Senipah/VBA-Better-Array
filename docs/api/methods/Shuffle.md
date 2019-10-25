---
title: Shuffle
parent: Methods
grand_parent: API
---

# Shuffle

## Description
The `Shuffle()` method Shuffles the order of elements in the array using the [Fisher–Yates algorithm](https://en.wikipedia.org/wiki/Fisher%E2%80%93Yates_shuffle#The_modern_algorithm) . 

## Syntax

*expression*.**Shuffle**([*Recurse*])

### Parameters

Name 
: `Recurse`

Type
: `Boolean`

Necessity
: Optional

Description
: If true and the array is jagged or multi-dimensional (which are stored as jagged internally), the order of all nested arrays will also be Shuffled.

### Returns

Type
: BetterArray `Object`

Description
: The current instance of the BetterArray object with the array's order Shuffled.


## Inspiration
* <http://www.ecma-international.org/ecma-262/10.0/index.html#sec-array.prototype.Shuffle>
* <https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/Shuffle>

[Back to Docs](https://senipah.github.io/VBA-Better-Array/)