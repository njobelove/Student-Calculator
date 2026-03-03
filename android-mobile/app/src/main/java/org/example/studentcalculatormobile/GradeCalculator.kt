package org.example.studentcalculatormobile

class GradeCalculator(private val passMark: Double = 60.0) {
    fun calculateAverage(scores: List<Double>): Double {
        return if (scores.isEmpty()) 0.0 else scores.sum() / scores.size
    }

    fun calculateGrade(average: Double): String {
        return when {
            average >= 90 -> "A"
            average >= 80 -> "B"
            average >= 70 -> "C"
            average >= passMark -> "D"
            else -> "F"
        }
    }
}
