# Task: Animal Class Hierarchy

## Objective
Practice inheritance, abstract classes, and polymorphism in Kotlin.

## Instructions
1. Create an abstract class `Animal` with:
   - A property `name` (passed via constructor).
   - An abstract `makeSound()` method that returns a `String`.
   - An abstract property `legs` (to be overridden by subclasses).

2. Create two concrete subclasses:
   - `Dog` – overrides `makeSound()` to return `"Woof!"` and sets `legs` appropriately.
   - `Cat` – overrides `makeSound()` to return `"Meow!"` and sets `legs` appropriately.

3. In the `main()` function:
   - Create a list containing one `Dog` and one `Cat` (use names of your choice).
   - Iterate over the list and print each animal's sound in the format:


## Hints
- Use `override` keyword for methods and properties.
- Remember that abstract classes cannot be instantiated directly.
- Polymorphism allows you to treat different subclasses as their common supertype.

## Bonus
- Add a third subclass (e.g., `Bird`) with a different number of legs and sound.