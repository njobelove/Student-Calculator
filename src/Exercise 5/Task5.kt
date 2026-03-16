// 1. Define the Logger interface
interface Logger {
    fun log(message: String)
}

// 2a. ConsoleLogger implementation – prints to console
class ConsoleLogger : Logger {
    override fun log(message: String) {
        println(message)
    }
}

// 2b. FileLogger implementation – simulates writing to a file
class FileLogger : Logger {
    override fun log(message: String) {
        println("File: $message")
    }
}

// 3. Application class that delegates logging to a Logger
//    using Kotlin's class delegation (by keyword)
class Application(logger: Logger) : Logger by logger

// Test the implementation
fun main() {
    val app = Application(ConsoleLogger())
    app.log("App started")          // prints to console: App started

    val fileApp = Application(FileLogger())
    fileApp.log("Error occurred")    // prints: File: Error occurred
}