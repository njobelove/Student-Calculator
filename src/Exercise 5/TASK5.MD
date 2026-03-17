# Task5.kt – Logger Using Delegation

This document provides the complete solution for **Exercise 2: Implement a Logger Using Delegation**. The task demonstrates Kotlin's interface and class delegation to create a flexible logging system where logging behaviour can be swapped easily.

## Overview

The solution consists of:
- A `Logger` interface defining the logging contract.
- Two implementations: `ConsoleLogger` (prints to console) and `FileLogger` (simulates writing to a file).
- An `Application` class that delegates logging to any `Logger` instance using Kotlin's `by` keyword.

## Complete Code (Task5.kt)

```kotlin
/**
 * Logger interface – defines the logging contract.
 */
interface Logger {
    fun log(message: String)
}

/**
 * ConsoleLogger – prints log messages directly to the console.
 */
class ConsoleLogger : Logger {
    override fun log(message: String) {
        println(message)
    }
}

/**
 * FileLogger – simulates writing to a file by prefixing the message.
 * In a real application, you would replace this with actual file I/O.
 */
class FileLogger : Logger {
    override fun log(message: String) {
        println("File: $message")
    }
}

/**
 * Application class that delegates logging to the provided logger.
 * The 'by' keyword automatically implements the Logger interface
 * by forwarding all calls to the 'logger' parameter.
 */
class Application(logger: Logger) : Logger by logger

/**
 * Test the implementation.
 */
fun main() {
    val app = Application(ConsoleLogger())
    app.log("App started")           // Output: App started

    val fileApp = Application(FileLogger())
    fileApp.log("Error occurred")     // Output: File: Error occurred
}