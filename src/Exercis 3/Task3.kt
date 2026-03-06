interface Drawable {
    fun draw()
}

class Circle(private val radius: Int) : Drawable {
    override fun draw() {
        println("Circle -> radius: $radius, area: %.2f".format(Math.PI * radius * radius))
    }
}

class Square(private val side: Int) : Drawable {
    override fun draw() {
        println("Square -> side: $side, area: ${side * side}")
    }
}

class Rectangle(private val width: Int, private val height: Int) : Drawable {
    override fun draw() {
        println("Rectangle -> width: $width, height: $height, area: ${width * height}")
    }
}

fun main() {
    val shapes: List<Drawable> = listOf(
        Circle(3),
        Square(4),
        Rectangle(5, 2)
    )

    println("=== Shapes Output ===")
    shapes.forEach { it.draw() }
}
