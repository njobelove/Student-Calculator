# Task4.kt – `maxOf` Function Documentation

This document describes the Kotlin implementation of a generic `maxOf` function that finds the maximum element in a list. It handles any type that implements `Comparable` and returns `null` for an empty list.

## Overview

The `maxOf` function is a generic utility that works with lists of any type `T` where `T` implements the `Comparable<T>` interface. It uses the `fold` operation to traverse the list while keeping track of the largest element seen so far. For an empty list, it gracefully returns `null` to indicate that no maximum exists.

## Complete Code

```kotlin
// Optional imports – Kotlin automatically imports these packages
import kotlin.collections.listOf
import kotlin.collections.emptyList
import kotlin.io.println

/**
 * Returns the maximum element in a list, or `null` if the list is empty.
 *
 * @param list A list of elements that implement [Comparable]
 * @return The maximum element, or `null` for an empty list
 */
fun <T : Comparable<T>> maxOf(list: List<T>): T? {
    return list.fold(null) { max: T?, element: T ->
        if (max == null || element > max) element else max
    }
}

fun main() {
    println(maxOf(listOf(3, 7, 2, 9)))                // 9
    println(maxOf(listOf("apple", "banana", "kiwi"))) // "kiwi"
    println(maxOf(emptyList<Int>()))                   // null
}