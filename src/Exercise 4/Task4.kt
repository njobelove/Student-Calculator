// Optional imports (Kotlin imports these automatically)
import kotlin.collections.listOf
import kotlin.collections.emptyList
import kotlin.io.println

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