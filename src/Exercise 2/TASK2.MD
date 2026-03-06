# Task 2: Sealed Class for Network States

Define a sealed class `NetworkState` that represents the different states of a network request:

- `Loading` – an object with no data.
- `Success` – a data class holding a `String` (the loaded data).
- `Error` – a data class holding a `String` (the error message).

Write a function `handleState(state: NetworkState)` that prints an appropriate message for each state:

- For `Loading` → print `"Loading..."`
- For `Success` → print `"Success: <data>"`
- For `Error` → print `"Error: <message>"`

The sealed class ensures that all possible states are handled exhaustively.

## Example Usage

```kotlin
val states = listOf(
    NetworkState.Loading,
    NetworkState.Success("User data loaded"),
    NetworkState.Error("Network timeout")
)

states.forEach { handleState(it) }