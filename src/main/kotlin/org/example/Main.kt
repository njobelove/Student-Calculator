package org.example

import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.BufferedReader
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.FileReader
import java.net.URI
import java.util.Scanner

/**
 * GradeCalculator
 *
 * Properties:
 * - passMark: minimum average score to pass.
 *
 * Structure:
 * - calculateGrade(average): returns letter grade from numeric average.
 * - calculateAverage(scores): returns average score from a list.
 * - createSampleScoresFile(path): creates an Excel file with a Scores sheet.
 * - processExcel(inputPath, outputPath): reads Scores, writes Grades, saves workbook.
 * - printSheet(workbook, sheetName): prints an Excel sheet to console.
 */
class GradeCalculator(private val passMark: Double = 60.0) {
    fun normalizePath(path: String): String = path.trim().trim('"')

    fun isHttpUrl(value: String): Boolean {
        val v = value.trim()
        return v.startsWith("http://", ignoreCase = true) || v.startsWith("https://", ignoreCase = true)
    }
    fun isCsvPath(value: String): Boolean = value.trim().endsWith(".csv", ignoreCase = true)

    fun calculateGrade(average: Double): String {
        return when {
            average >= 90 -> "A"
            average >= 80 -> "B"
            average >= 70 -> "C"
            average >= passMark -> "D"
            else -> "F"
        }
    }

    fun calculateAverage(scores: List<Double>): Double {
        return if (scores.isEmpty()) 0.0 else scores.sum() / scores.size
    }

    fun createSampleScoresFile(path: String) {
        val file = File(path)
        if (file.exists()) return

        XSSFWorkbook().use { workbook ->
            val scoresSheet = workbook.createSheet("Scores")
            val header = scoresSheet.createRow(0)
            header.createCell(0).setCellValue("Student")
            header.createCell(1).setCellValue("Math")
            header.createCell(2).setCellValue("Science")
            header.createCell(3).setCellValue("English")

            val sampleRows = listOf(
                listOf("Alice", "88", "91", "84"),
                listOf("Bob", "72", "68", "75"),
                listOf("Charlie", "95", "93", "97")
            )

            sampleRows.forEachIndexed { index, rowValues ->
                val row = scoresSheet.createRow(index + 1)
                rowValues.forEachIndexed { col, value ->
                    if (col == 0) {
                        row.createCell(col).setCellValue(value)
                    } else {
                        row.createCell(col).setCellValue(value.toDouble())
                    }
                }
            }

            FileOutputStream(file).use { fos -> workbook.write(fos) }
        }
    }

    fun processExcel(inputPath: String, outputPath: String) {
        val inputFile = File(inputPath)
        if (!inputFile.exists()) {
            println("Input file not found: $inputPath")
            return
        }
        println("Using input workbook: ${inputFile.absolutePath}")

        try {
            FileInputStream(inputFile).use { fis ->
                XSSFWorkbook(fis).use { workbook ->
                    val sheetNames = (0 until workbook.numberOfSheets).map { workbook.getSheetAt(it).sheetName }
                    println("Workbook sheets: ${sheetNames.joinToString(", ")}")
                    val scoresSheet = workbook.getSheet("Scores") ?: workbook.getSheetAt(0)
                    val inputSheetName = scoresSheet.sheetName

                    val existingGrades = workbook.getSheet("Grades")
                    if (existingGrades != null) {
                        val index = workbook.getSheetIndex(existingGrades)
                        workbook.removeSheetAt(index)
                    }
                    val gradesSheet = workbook.createSheet("Grades")

                    val headerRows = (0..2).mapNotNull { idx -> scoresSheet.getRow(idx) }
                    val coursesStart = headerRows
                        .map { row ->
                            row.firstOrNull { cell ->
                                cell?.toString()?.trim()?.equals("courses", ignoreCase = true) == true
                            }?.columnIndex ?: -1
                        }
                        .firstOrNull { it >= 0 } ?: -1

                    val courseHeaderRow = headerRows
                        .maxByOrNull { row ->
                            row.count { cell ->
                                val v = cell?.toString()?.trim()?.lowercase().orEmpty()
                                v.isNotEmpty() && v != "courses" && v != "ca" && v != "exam"
                            }
                        }

                    val courseColumns = linkedMapOf<String, MutableList<Int>>()
                    if (coursesStart >= 0 && courseHeaderRow != null) {
                        val maxCols = courseHeaderRow.lastCellNum.toInt().coerceAtLeast(0)
                        var currentCourse = ""
                        for (col in coursesStart + 1 until maxCols) {
                            val headerValue = courseHeaderRow.getCell(col)?.toString()?.trim().orEmpty()
                            if (headerValue.isNotBlank()) {
                                val v = headerValue.lowercase()
                                if (v != "courses" && v != "ca" && v != "exam") {
                                    currentCourse = headerValue
                                }
                            }
                            if (currentCourse.isNotBlank()) {
                                courseColumns.getOrPut(currentCourse) { mutableListOf() }.add(col)
                            }
                        }
                    }

                    val header = gradesSheet.createRow(0)
                    if (courseColumns.isNotEmpty()) {
                        var colIndex = 0
                        header.createCell(colIndex++).setCellValue("Student")
                        for (course in courseColumns.keys) {
                            header.createCell(colIndex++).setCellValue("$course Total")
                            header.createCell(colIndex++).setCellValue("$course Grade")
                        }
                    } else {
                        header.createCell(0).setCellValue("Student")
                        header.createCell(1).setCellValue("Average")
                        header.createCell(2).setCellValue("Grade")
                    }

                    val nameIndex = 1
                    var outRowIndex = 1
                    for (i in 1..scoresSheet.lastRowNum) {
                        val row = scoresSheet.getRow(i) ?: continue
                        val studentName = row.getCell(nameIndex)?.toString()?.trim().orEmpty()
                        if (studentName.isBlank()) continue

                        if (courseColumns.isNotEmpty()) {
                            val outRow = gradesSheet.createRow(outRowIndex++)
                            var colIndex = 0
                            outRow.createCell(colIndex++).setCellValue(studentName)
                            var hasAnyScore = false
                            for ((_, cols) in courseColumns) {
                                var total = 0.0
                                var count = 0
                                for (col in cols) {
                                    val cell = row.getCell(col) ?: continue
                                    when (cell.cellType) {
                                        CellType.NUMERIC -> {
                                            total += cell.numericCellValue
                                            count += 1
                                        }
                                        CellType.STRING -> {
                                            val v = cell.stringCellValue.toDoubleOrNull()
                                            if (v != null) {
                                                total += v
                                                count += 1
                                            }
                                        }
                                        else -> {}
                                    }
                                }
                                if (count > 0) {
                                    hasAnyScore = true
                                    outRow.createCell(colIndex++).setCellValue(String.format("%.2f", total))
                                    outRow.createCell(colIndex++).setCellValue(calculateGrade(total))
                                } else {
                                    outRow.createCell(colIndex++).setCellValue("")
                                    outRow.createCell(colIndex++).setCellValue("")
                                }
                            }
                            if (!hasAnyScore) {
                                gradesSheet.removeRow(outRow)
                                outRowIndex -= 1
                            }
                        } else {
                            val scores = mutableListOf<Double>()
                            for (col in 1 until row.lastCellNum) {
                                val cell = row.getCell(col) ?: continue
                                when (cell.cellType) {
                                    CellType.NUMERIC -> scores.add(cell.numericCellValue)
                                    CellType.STRING -> cell.stringCellValue.toDoubleOrNull()?.let { scores.add(it) }
                                    else -> {}
                                }
                            }
                            if (scores.isEmpty()) continue

                            val average = calculateAverage(scores)
                            val grade = calculateGrade(average)

                            val outRow = gradesSheet.createRow(outRowIndex++)
                            outRow.createCell(0).setCellValue(studentName)
                            outRow.createCell(1).setCellValue(String.format("%.2f", average))
                            outRow.createCell(2).setCellValue(grade)
                        }
                    }

                    val outputFile = File(outputPath)
                    if (outputFile.parentFile != null) outputFile.parentFile.mkdirs()
                    FileOutputStream(outputFile).use { fos -> workbook.write(fos) }

                    println("Saved output workbook: ${outputFile.absolutePath}")
                    printSheet(workbook, inputSheetName)
                    printSheet(workbook, "Grades")
                }
            }
        } catch (ex: Exception) {
            println("Failed to process Excel: ${ex.message}")
            val asCsv = looksLikeCsv(inputFile)
            if (asCsv) {
                val csvOut = if (outputPath.endsWith(".xlsx", ignoreCase = true)) {
                    outputPath.dropLast(5) + ".csv"
                } else {
                    outputPath + ".csv"
                }
                println("File does not look like a valid .xlsx. Trying CSV parser instead.")
                processCsv(inputFile.absolutePath, csvOut)
            } else {
                ex.printStackTrace()
            }
        }
    }

    private fun looksLikeCsv(file: File): Boolean {
        return try {
            val bytes = file.inputStream().use { it.readNBytes(1024) }
            if (bytes.isEmpty()) return false
            if (bytes.any { it == 0.toByte() }) return false
            val text = bytes.toString(Charsets.UTF_8)
            text.contains(",") || text.contains(";")
        } catch (_: Exception) {
            false
        }
    }

    private fun parseCsvLine(line: String): List<String> {
        val result = mutableListOf<String>()
        val sb = StringBuilder()
        var inQuotes = false
        var i = 0
        while (i < line.length) {
            val ch = line[i]
            if (ch == '"') {
                if (inQuotes && i + 1 < line.length && line[i + 1] == '"') {
                    sb.append('"')
                    i += 1
                } else {
                    inQuotes = !inQuotes
                }
            } else if (ch == ',' && !inQuotes) {
                result.add(sb.toString())
                sb.setLength(0)
            } else {
                sb.append(ch)
            }
            i += 1
        }
        result.add(sb.toString())
        return result
    }

    fun processCsv(inputPath: String, outputPath: String) {
        val inputFile = File(inputPath)
        if (!inputFile.exists()) {
            println("Input file not found: $inputPath")
            return
        }
        println("Using input CSV: ${inputFile.absolutePath}")

        val rows = mutableListOf<List<String>>()
        BufferedReader(FileReader(inputFile)).use { reader ->
            var line = reader.readLine()
            while (line != null) {
                if (line.isNotBlank()) rows.add(parseCsvLine(line))
                line = reader.readLine()
            }
        }

        if (rows.isEmpty()) {
            println("CSV is empty.")
            return
        }

        val headerRows = rows.take(3)
        val nameIndex = 1
        var coursesRowIndex = -1
        var coursesStart = -1
        for ((idx, row) in headerRows.withIndex()) {
            val col = row.indexOfFirst { it.trim().equals("courses", ignoreCase = true) }
            if (col >= 0) {
                coursesRowIndex = idx
                coursesStart = col
                break
            }
        }

        val scoreIndices = if (coursesStart >= 0) {
            val maxCols = headerRows.maxOfOrNull { it.size } ?: 0
            val indices = mutableSetOf<Int>()
            for (col in coursesStart + 1 until maxCols) {
                val hasHeaderText = headerRows.any { row -> col < row.size && row[col].trim().isNotEmpty() }
                if (hasHeaderText) indices.add(col)
            }
            indices
        } else {
            emptySet()
        }

        val courseHeaderRow = if (coursesRowIndex >= 0) {
            headerRows.drop(coursesRowIndex + 1).firstOrNull { row ->
                row.any {
                    val v = it.trim().lowercase()
                    v.isNotEmpty() && v != "courses" && v != "ca" && v != "exam"
                }
            } ?: headerRows.firstOrNull().orEmpty()
        } else {
            headerRows.firstOrNull().orEmpty()
        }

        val courseColumns = linkedMapOf<String, MutableList<Int>>()
        if (coursesStart >= 0) {
            val maxCols = headerRows.maxOfOrNull { it.size } ?: 0
            var currentCourse = ""
            for (col in coursesStart + 1 until maxCols) {
                val headerValue = if (col < courseHeaderRow.size) courseHeaderRow[col].trim() else ""
                if (headerValue.isNotBlank()) {
                    val v = headerValue.lowercase()
                    if (v != "courses" && v != "ca" && v != "exam") {
                        currentCourse = headerValue
                    }
                }
                if (currentCourse.isNotBlank()) {
                    courseColumns.getOrPut(currentCourse) { mutableListOf() }.add(col)
                }
            }
        }

        val gradesRows = mutableListOf<List<String>>()
        if (courseColumns.isNotEmpty()) {
            val header = mutableListOf("Student")
            for (course in courseColumns.keys) {
                header.add("$course Total")
                header.add("$course Grade")
            }
            gradesRows.add(header)
        } else {
            gradesRows.add(listOf("Student", "Average", "Grade"))
        }

        for (row in rows.drop(1)) {
            if (row.isEmpty()) continue
            val studentName = if (nameIndex < row.size) row[nameIndex].trim() else ""
            if (studentName.isBlank()) continue

            if (courseColumns.isNotEmpty()) {
                val out = mutableListOf(studentName)
                var hasAnyScore = false
                for ((course, cols) in courseColumns) {
                    var total = 0.0
                    var count = 0
                    for (idx in cols) {
                        if (idx >= row.size) continue
                        val v = row[idx].trim()
                        if (v.isEmpty()) continue
                        val num = v.toDoubleOrNull() ?: continue
                        total += num
                        count += 1
                    }
                    if (count > 0) {
                        hasAnyScore = true
                        out.add(String.format("%.2f", total))
                        out.add(calculateGrade(total))
                    } else {
                        out.add("")
                        out.add("")
                    }
                }
                if (hasAnyScore) gradesRows.add(out)
            } else {
                val scores = mutableListOf<Double>()
                if (scoreIndices.isNotEmpty()) {
                    for (idx in scoreIndices) {
                        if (idx >= row.size) continue
                        val v = row[idx].trim()
                        if (v.isEmpty()) continue
                        v.toDoubleOrNull()?.let { scores.add(it) }
                    }
                } else {
                    for ((idx, cell) in row.withIndex()) {
                        if (idx <= nameIndex) continue
                        val v = cell.trim()
                        if (v.isEmpty()) continue
                        v.toDoubleOrNull()?.let { scores.add(it) }
                    }
                }
                if (scores.isEmpty()) continue

                val average = calculateAverage(scores)
                val grade = calculateGrade(average)
                gradesRows.add(listOf(studentName, String.format("%.2f", average), grade))
            }
        }

        val outputFile = File(outputPath)
        if (outputFile.parentFile != null) outputFile.parentFile.mkdirs()
        FileOutputStream(outputFile).use { fos ->
            val outLines = gradesRows.joinToString("\n") { r ->
                r.joinToString(",") { it.replace("\"", "\"\"").let { v -> if (v.contains(",") || v.contains("\"")) "\"$v\"" else v } }
            }
            fos.write(outLines.toByteArray())
        }

        println("Saved output CSV: ${outputFile.absolutePath}")
        printCsv("Scores (CSV)", rows)
        printCsv("Grades (CSV)", gradesRows)
    }

    private fun printCsv(title: String, rows: List<List<String>>) {
        println("\n=== $title ===")
        for (row in rows) {
            println(row.joinToString(" | ") { it.trim() })
        }
    }

    fun downloadExcelFromUrl(url: String, destinationPath: String): Boolean {
        return try {
            URI(url).toURL().openStream().use { input ->
                FileOutputStream(destinationPath).use { output ->
                    input.copyTo(output)
                }
            }
            true
        } catch (ex: Exception) {
            println("Failed to download Excel from URL: ${ex.message}")
            false
        }
    }

    fun printSheet(workbook: XSSFWorkbook, sheetName: String) {
        val sheet = workbook.getSheet(sheetName)
        if (sheet == null) {
            println("Sheet '$sheetName' not found.")
            return
        }

        println("\n=== $sheetName Sheet ===")
        var printed = false
        for (i in 0..sheet.lastRowNum) {
            val row = sheet.getRow(i) ?: continue
            val values = (0 until row.lastCellNum).map { col ->
                val cell = row.getCell(col)
                when (cell?.cellType) {
                    CellType.NUMERIC -> String.format("%.2f", cell.numericCellValue)
                    CellType.STRING -> cell.stringCellValue
                    CellType.BOOLEAN -> cell.booleanCellValue.toString()
                    else -> ""
                }
            }
            println(values.joinToString(" | "))
            printed = true
        }
        if (!printed) {
            println("(empty sheet)")
        }
    }
}

/**
 * Main procedure:
 * 1. Choose console mode or Excel mode.
 * 2. Console mode asks student name and scores, then prints average and grade.
 * 3. Excel mode creates sample Scores file (if missing), reads it, creates Grades sheet,
 *    saves new workbook, and prints both sheets to the console.
 */
fun main() {
    val scanner = Scanner(System.`in`)
    val calculator = GradeCalculator()
    val projectRoot = File(".").absoluteFile

    println("--- Student Grade Calculator (Kotlin Console App) ---")
    println("1. Grade from console scores")
    println("2. Grade from local Excel file (Scores -> Grades)")
    println("3. Open mobile version instructions")
    println("4. Grade from online Excel URL (Scores -> Grades)")
    print("Select option (1, 2, 3 or 4): ")

    val option = if (scanner.hasNextLine()) {
        scanner.nextLine().trim().toIntOrNull()
    } else {
        2
    }
    when (option) {
        1 -> {
            print("Student name: ")
            val student = scanner.nextLine().trim()
            print("How many scores? ")
            val count = scanner.nextLine().trim().toIntOrNull() ?: 0

            if (count <= 0) {
                println("Invalid number of scores.")
                return
            }

            val scores = mutableListOf<Double>()
            for (i in 1..count) {
                print("Enter score $i: ")
                val score = scanner.nextLine().trim().toDoubleOrNull()
                if (score == null) {
                    println("Invalid score input. Stopping.")
                    return
                }
                scores.add(score)
            }

            val avg = calculator.calculateAverage(scores)
            val grade = calculator.calculateGrade(avg)
            println("$student -> Average: ${"%.2f".format(avg)}, Grade: $grade")
        }
        2 -> {
            print("Enter local Excel file path (press Enter for scores.xlsx): ")
            val enteredPath = if (scanner.hasNextLine()) scanner.nextLine().trim() else ""
            val inputFile = if (enteredPath.isBlank()) {
                val cwd = File(".")
                val candidates = cwd.listFiles { file ->
                    file.isFile && file.name.endsWith(".xlsx", ignoreCase = true) &&
                        !file.name.equals("student_grades.xlsx", ignoreCase = true) &&
                        !file.name.equals("student_grades_from_online.xlsx", ignoreCase = true)
                }?.sortedBy { it.name.lowercase() }.orEmpty()

                if (File("scores.xlsx").exists()) {
                    "scores.xlsx"
                } else if (candidates.isNotEmpty()) {
                    candidates.first().name
                } else {
                    "scores.xlsx"
                }
            } else calculator.normalizePath(enteredPath)
            if (calculator.isCsvPath(inputFile)) {
                val outputFile = File(projectRoot, "student_grades.csv").absolutePath
                calculator.processCsv(inputFile, outputFile)
                println("Input file : ${File(inputFile).absolutePath}")
                println("Output file: ${File(outputFile).absolutePath}")
            } else {
                val outputFile = File(projectRoot, "student_grades.xlsx").absolutePath
                if (inputFile == "scores.xlsx") {
                    calculator.createSampleScoresFile(inputFile)
                }
                calculator.processExcel(inputFile, outputFile)
                println("Input workbook : ${File(inputFile).absolutePath}")
                println("Output workbook: ${File(outputFile).absolutePath}")
            }
        }

        3 -> {
            println("Mobile app project folder:")
            println("C:\\Users\\user\\Desktop\\StudentCalculator\\StudentCalculator\\android-mobile")
            println("Open that folder in Android Studio and run the app on emulator/device.")
        }

        4 -> {
            print("Paste downloadable Excel URL (or a local path): ")
            val source = if (scanner.hasNextLine()) scanner.nextLine().trim() else ""
            if (source.isBlank()) {
                println("No source provided.")
                return
            }

            if (!calculator.isHttpUrl(source)) {
                val localPath = calculator.normalizePath(source)
                if (calculator.isCsvPath(localPath)) {
                    val outputFile = File(projectRoot, "student_grades_from_online.csv").absolutePath
                    calculator.processCsv(localPath, outputFile)
                    println("Input file : ${File(localPath).absolutePath}")
                    println("Output file: ${File(outputFile).absolutePath}")
                } else {
                    val outputFile = File(projectRoot, "student_grades_from_online.xlsx").absolutePath
                    calculator.processExcel(localPath, outputFile)
                    println("Input workbook : ${File(localPath).absolutePath}")
                    println("Output workbook: ${File(outputFile).absolutePath}")
                }
            } else {
                if (calculator.isCsvPath(source)) {
                    val downloadedInput = "scores_online.csv"
                    val downloaded = calculator.downloadExcelFromUrl(source, downloadedInput)
                    if (!downloaded) return

                    val outputFile = File(projectRoot, "student_grades_from_online.csv").absolutePath
                    calculator.processCsv(downloadedInput, outputFile)
                    println("Downloaded file: ${File(downloadedInput).absolutePath}")
                    println("Output file    : ${File(outputFile).absolutePath}")
                } else {
                    val downloadedInput = "scores_online.xlsx"
                    val downloaded = calculator.downloadExcelFromUrl(source, downloadedInput)
                    if (!downloaded) return

                    val outputFile = File(projectRoot, "student_grades_from_online.xlsx").absolutePath
                    calculator.processExcel(downloadedInput, outputFile)
                    println("Downloaded workbook: ${File(downloadedInput).absolutePath}")
                    println("Output workbook    : ${File(outputFile).absolutePath}")
                }
            }
        }

        else -> println("Invalid option. Please run again and choose 1, 2, 3 or 4.")
    }
}
