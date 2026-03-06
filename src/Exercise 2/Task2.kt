sealed class NetworkState {
    object Loading : NetworkState()
    data class Success(val data: String) : NetworkState()
    data class Error(val message: String) : NetworkState()
}

// Function to handle each state
fun handleState(state: NetworkState) {
    when (state) {
        NetworkState.Loading -> println("Loading...")
        is NetworkState.Success -> println("Success: ${state.data}")
        is NetworkState.Error -> println("Error: ${state.message}")
    }
    // No 'else' needed because the when is exhaustive for sealed classes
}

// Example usage
fun main() {
    val states = listOf(
        NetworkState.Loading,
        NetworkState.Success("User data loaded"),
        NetworkState.Error("Network timeout")
    )
    states.forEach { handleState(it) }
}