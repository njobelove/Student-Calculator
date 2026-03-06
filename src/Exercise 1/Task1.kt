abstract class Animal(val name: String) {
    abstract fun makeSound(): String
    abstract val legs: Int   // abstract property, each subclass must override
}

// Dog subclass
class Dog(name: String) : Animal(name) {
    override val legs = 4
    override fun makeSound() = "Woof!"
}

// Cat subclass
class Cat(name: String) : Animal(name) {
    override val legs = 4
    override fun makeSound() = "Meow!"
}

fun main() {
    // Create a list of animals
    val animals = listOf(Dog("Buddy"), Cat("Whiskers"))

    // Iterate and print each sound along with name
    for (animal in animals) {
        println("${animal.name} says ${animal.makeSound()}")
    }
}
