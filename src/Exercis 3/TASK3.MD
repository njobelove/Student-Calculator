# Task: Drawable Interface with Circle and Square

Define an interface `Drawable` that declares a function `draw()`.  
Create two classes `Circle` and `Square` that implement this interface. Each class should have appropriate properties (`radius` for `Circle`, `sideLength` for `Square`) and override the `draw()` function to print a simple ASCII representation of the shape.

## Requirements

- The `Drawable` interface must contain the function `draw()`.
- `Circle` should have a `radius: Int` property.
- `Square` should have a `sideLength: Int` property.
- The `draw()` method should output an ASCII drawing to the console.  
  (You can design simple representations; for example, a small circle might look like a ring, and a square like a hollow box.)

## Example Implementation Idea

```kotlin
interface Drawable {
    fun draw()
}

class Circle(val radius: Int) : Drawable {
    override fun draw() {
        // Simple ASCII circle (size depends on radius? Here fixed for simplicity)
        println("  ***  ")
        println(" *   * ")
        println("*     *")
        println(" *   * ")
        println("  ***  ")
    }
}

class Square(val sideLength: Int) : Drawable {
    override fun draw() {
        // Simple hollow square
        println("****")
        println("*  *")
        println("*  *")
        println("****")
    }
}