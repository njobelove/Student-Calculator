package org.example

import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.awt.BorderLayout
import java.awt.FlowLayout
import java.awt.Font
import java.awt.GraphicsEnvironment
import java.io.*
import java.net.URI
import java.util.Scanner
import javax.swing.*
import javax.swing.table.DefaultTableModel

// ========================== Output Format Enum ==========================
enum class OutputFormat { EXCEL, PDF, HTML }

// ========================== Interfaces ==========================
interface FileProcessor {
    fun processFile(file: File): Any?   // returns processed data (e.g., List<StudentResult>)
}

interface OutputGenerator {
    fun generateOutput(data: Any, format: OutputFormat): File
}

interface Downloadable {
    fun download(url: String): File
}

// ========================== Abstract Calculator ==========================
abstract class Calculator(protected val passMark: Double = 60.0) {
    fun calculateGrade(average: Double): String = when {
        average >= 90 -> "A"
        average >= 80 -> "B"
        average >= 70 -> "C"
        average >= passMark -> "D"
        else -> "F"
    }

    fun calculateAverage(scores: List<Double>): Double =
        if (scores.isEmpty()) 0.0 else scores.sum() / scores.size

    // Higher‑order function: computes average using a passed aggregator
    fun computeAverage(scores: List<Double>, aggregator: (List<Double>) -> Double): Double = aggregator(scores)

    fun processScores(scores: List<Double>, aggregator: (List<Double>) -> Double): Pair<Double, String> {
        val avg = computeAverage(scores, aggregator)
        return avg to calculateGrade(avg)
    }

    // Higher‑order function with collection operator (map)
    fun processScoresWithOperation(
        scores: List<Double>,
        operation: (List<Double>) -> List<Double>,
        aggregator: (List<Double>) -> Double
    ): Double {
        val transformed = operation(scores)
        return aggregator(transformed)
    }
}

// ========================== GradeCalculator (Child Class) ==========================
class GradeCalculator(passMark: Double = 60.0) : Calculator(passMark),
    FileProcessor, OutputGenerator, Downloadable {

    // ---------- Downloadable ----------
    override fun download(url: String): File {
        val fileName = File(URI(url).path).name.ifBlank { "downloaded.tmp" }
        val tempFile = File(fileName)
        println("Downloading from $url ...")
        URI(url).toURL().openStream().use { input ->
            FileOutputStream(tempFile).use { output -> input.copyTo(output) }
        }
        println("Downloaded to ${tempFile.absolutePath}")
        return tempFile
    }

    // ---------- FileProcessor ----------
    override fun processFile(file: File): Any? {
        return when {
            file.extension.equals("xlsx", ignoreCase = true) -> processExcel(file)
            file.extension.equals("csv", ignoreCase = true) -> processCsv(file)
            file.extension.equals("pdf", ignoreCase = true) -> processPdf(file)
            file.extension.equals("html", ignoreCase = true) || file.extension.equals("htm", ignoreCase = true) ->
                processHtml(file)
            else -> {
                println("Unsupported file type: ${file.extension}")
                null
            }
        }
    }

    // ---------- OutputGenerator ----------
    override fun generateOutput(data: Any, format: OutputFormat): File {
        return when (format) {
            OutputFormat.EXCEL -> generateExcelOutput(data)
            OutputFormat.PDF -> generatePdfOutput(data)
            OutputFormat.HTML -> generateHtmlOutput(data)
        }
    }

    // ---------- Private processing methods ----------
    private fun processExcel(file: File): List<StudentResult> {
        println("Processing Excel file: ${file.absolutePath}")
        val results = mutableListOf<StudentResult>()
        FileInputStream(file).use { fis ->
            XSSFWorkbook(fis).use { workbook ->
                val scoresSheet = workbook.getSheet("Scores") ?: workbook.getSheetAt(0)
                val inputSheetName = scoresSheet.sheetName

                workbook.getSheet("Grades")?.let {
                    workbook.removeSheetAt(workbook.getSheetIndex(it))
                }
                val gradesSheet = workbook.createSheet("Grades")

                val layout = buildExcelLayout(scoresSheet)
                val nameIndex = layout.nameIndex
                val courseColumns = layout.courseColumns

                // Grades header
                val header = gradesSheet.createRow(0)
                var colIdx = 0
                header.createCell(colIdx++).setCellValue("Student")
                for (course in courseColumns.keys) {
                    header.createCell(colIdx++).setCellValue("$course Total")
                    header.createCell(colIdx++).setCellValue("$course Grade")
                }

                var outRowIdx = 1
                for (i in layout.dataStartRow..scoresSheet.lastRowNum) {
                    val row = scoresSheet.getRow(i) ?: continue
                    val student = row.getCell(nameIndex)?.toString()?.trim().orEmpty()
                    if (student.isBlank()) continue

                    val outRow = gradesSheet.createRow(outRowIdx++)
                    var outCol = 0
                    outRow.createCell(outCol++).setCellValue(student)

                    val studentTotals = mutableMapOf<String, Double>()
                    var hasAny = false
                    for ((course, cols) in courseColumns) {
                        val (primaryCol, secondaryCol) = cols
                        val primary = getNumericCellValue(row.getCell(primaryCol))
                        val total = if (secondaryCol == null) {
                            primary
                        } else {
                            val secondary = getNumericCellValue(row.getCell(secondaryCol))
                            if (primary != null && secondary != null) primary + secondary else null
                        }
                        if (total != null) {
                            hasAny = true
                            studentTotals[course] = total
                            outRow.createCell(outCol++).setCellValue(total)
                            outRow.createCell(outCol++).setCellValue(calculateGrade(total))
                        } else {
                            outRow.createCell(outCol++).setCellValue("")
                            outRow.createCell(outCol++).setCellValue("")
                        }
                    }
                    if (hasAny) results.add(StudentResult(student, studentTotals))
                    else {
                        gradesSheet.removeRow(outRow)
                        outRowIdx--
                    }
                }

                // Save workbook
                val outputFile = File("student_grades.xlsx")
                FileOutputStream(outputFile).use { fos -> workbook.write(fos) }
                println("Saved: ${outputFile.absolutePath}")

                // Return both results and the workbook (for GUI display)
                // We'll store the workbook in a temporary map or return it as part of a custom object.
                // To keep it simple, we'll just return the results and rely on the GUI to reload the file.
                // The GUI will call a helper to display the workbook.
            }
        }
        return results
    }

    private fun processCsv(file: File): List<StudentResult> {
        println("Processing CSV file: ${file.absolutePath}")
        val rows = mutableListOf<List<String>>()
        BufferedReader(FileReader(file)).use { reader ->
            var line = reader.readLine()
            while (line != null) {
                if (line.isNotBlank()) rows.add(parseCsvLine(line))
                line = reader.readLine()
            }
        }
        if (rows.isEmpty()) return emptyList()

        val workbook = XSSFWorkbook()
        val scoresSheet = workbook.createSheet("Scores")
        rows.forEachIndexed { r, rowData ->
            val excelRow = scoresSheet.createRow(r)
            rowData.forEachIndexed { c, cellValue ->
                excelRow.createCell(c).setCellValue(cellValue)
            }
        }

        val layout = buildCsvLayout(rows)
        val nameIndex = layout.nameIndex
        val courseColumns = layout.courseColumns
        val dataRows = rows.drop(layout.dataStartRow)

        val gradesSheet = workbook.createSheet("Grades")
        val header = gradesSheet.createRow(0)
        var colIdx = 0
        header.createCell(colIdx++).setCellValue("Student")
        for (course in courseColumns.keys) {
            header.createCell(colIdx++).setCellValue("$course Total")
            header.createCell(colIdx++).setCellValue("$course Grade")
        }

        val results = mutableListOf<StudentResult>()
        var outRowIdx = 1
        for (row in dataRows) {
            val student = row.getOrNull(nameIndex)?.trim().orEmpty()
            if (student.isBlank()) continue

            val outRow = gradesSheet.createRow(outRowIdx++)
            var outCol = 0
            outRow.createCell(outCol++).setCellValue(student)

            val studentTotals = mutableMapOf<String, Double>()
            var hasAny = false
            for ((course, cols) in courseColumns) {
                val (primaryCol, secondaryCol) = cols
                val primary = row.getOrNull(primaryCol)?.trim()?.toDoubleOrNull()
                val total = if (secondaryCol == null) {
                    primary
                } else {
                    val secondary = row.getOrNull(secondaryCol)?.trim()?.toDoubleOrNull()
                    if (primary != null && secondary != null) primary + secondary else null
                }
                if (total != null) {
                    hasAny = true
                    studentTotals[course] = total
                    outRow.createCell(outCol++).setCellValue(total)
                    outRow.createCell(outCol++).setCellValue(calculateGrade(total))
                } else {
                    outRow.createCell(outCol++).setCellValue("")
                    outRow.createCell(outCol++).setCellValue("")
                }
            }
            if (hasAny) results.add(StudentResult(student, studentTotals))
            else {
                gradesSheet.removeRow(outRow)
                outRowIdx--
            }
        }

        val outputFile = File("student_grades.xlsx")
        FileOutputStream(outputFile).use { fos -> workbook.write(fos) }
        println("Saved: ${outputFile.absolutePath}")
        workbook.close()
        return results
    }

    private fun processPdf(file: File): List<StudentResult> {
        println("PDF processing not fully implemented – returning empty results.")
        return emptyList()
    }

    private fun processHtml(file: File): List<StudentResult> {
        println("HTML processing not fully implemented – returning empty results.")
        return emptyList()
    }

    // ---------- Output generation (manual entry only) ----------
    private fun generateExcelOutput(data: Any): File {
        val results = data as List<StudentResult>
        val workbook = XSSFWorkbook()
        val sheet = workbook.createSheet("Grades")
        val header = sheet.createRow(0)
        var col = 0
        header.createCell(col++).setCellValue("Student")
        if (results.isNotEmpty()) {
            val courses = results.first().courseTotals.keys.sorted()
            courses.forEach { course ->
                header.createCell(col++).setCellValue("$course Total")
                header.createCell(col++).setCellValue("$course Grade")
            }
        }
        results.forEachIndexed { idx, res ->
            val row = sheet.createRow(idx + 1)
            var c = 0
            row.createCell(c++).setCellValue(res.student)
            res.courseTotals.toSortedMap().forEach { (_, total) ->
                row.createCell(c++).setCellValue(total)
                row.createCell(c++).setCellValue(calculateGrade(total))
            }
        }
        val outFile = File("student_grades.xlsx")
        FileOutputStream(outFile).use { workbook.write(it) }
        return outFile
    }

    private fun generatePdfOutput(data: Any): File {
        println("PDF generation placeholder – creating dummy file.")
        val outFile = File("student_grades.pdf")
        outFile.writeText("PDF output would be generated here.")
        return outFile
    }

    private fun generateHtmlOutput(data: Any): File {
        println("HTML generation placeholder – creating dummy file.")
        val outFile = File("student_grades.html")
        val html = buildString {
            appendLine("<html><body><table border='1'>")
            appendLine("<tr><th>Student</th><th>Course</th><th>Total</th><th>Grade</th></tr>")
            (data as List<StudentResult>).forEach { res ->
                res.courseTotals.forEach { (course, total) ->
                    appendLine("<tr><td>${res.student}</td><td>$course</td><td>$total</td><td>${calculateGrade(total)}</td></tr>")
                }
            }
            appendLine("</table></body></html>")
        }
        outFile.writeText(html)
        return outFile
    }

    // ---------- Helpers (same as before) ----------
    private data class ParsedLayout(
        val nameIndex: Int,
        val dataStartRow: Int,
        val courseColumns: LinkedHashMap<String, Pair<Int, Int?>>
    )

    private fun buildExcelLayout(sheet: org.apache.poi.ss.usermodel.Sheet): ParsedLayout {
        val nameIndex = detectNameIndex(sheet)
        val coursesMarker = detectCoursesMarkerFromSheet(sheet)

        if (coursesMarker != null) {
            val (coursesRowIndex, coursesStartCol) = coursesMarker
            val courseNameRow = sheet.getRow((coursesRowIndex + 1).coerceAtMost(sheet.lastRowNum))
            val subHeaderRow = sheet.getRow((coursesRowIndex + 2).coerceAtMost(sheet.lastRowNum))
            val hasCaExam = subHeaderRow != null && (coursesStartCol until subHeaderRow.lastCellNum).any { col ->
                val value = subHeaderRow.getCell(col)?.toString()?.trim()?.lowercase().orEmpty()
                value == "ca" || value == "exam"
            }

            val columns = LinkedHashMap<String, Pair<Int, Int?>>()
            var currentCourse = ""
            val maxCol = listOf(
                courseNameRow?.lastCellNum?.toInt() ?: 0,
                subHeaderRow?.lastCellNum?.toInt() ?: 0
            ).maxOrNull() ?: 0

            for (col in coursesStartCol until maxCol) {
                val candidate = courseNameRow?.getCell(col)?.toString()?.trim().orEmpty()
                if (candidate.isNotBlank() && candidate.lowercase() !in setOf("courses", "ca", "exam")) {
                    currentCourse = candidate
                    if (!hasCaExam && col != nameIndex) {
                        columns[currentCourse] = col to null
                    }
                }
                if (!hasCaExam || currentCourse.isBlank()) continue

                when (subHeaderRow?.getCell(col)?.toString()?.trim()?.lowercase()) {
                    "ca" -> {
                        val exam = columns[currentCourse]?.second
                        columns[currentCourse] = col to exam
                    }
                    "exam" -> {
                        val ca = columns[currentCourse]?.first ?: -1
                        columns[currentCourse] = ca to col
                    }
                }
            }

            if (hasCaExam) {
                columns.entries.removeIf { (_, cols) ->
                    val (caCol, examCol) = cols
                    caCol < 0 || examCol == null || examCol < 0
                }
            }

            if (columns.isNotEmpty()) {
                val dataStart = if (hasCaExam) coursesRowIndex + 3 else coursesRowIndex + 2
                return ParsedLayout(nameIndex, dataStart, columns)
            }
        }

        val headerRowIndex = (0..sheet.lastRowNum).firstOrNull { r ->
            val row = sheet.getRow(r) ?: return@firstOrNull false
            var textCells = 0
            for (c in 0 until row.lastCellNum) {
                if (row.getCell(c)?.toString()?.trim().isNullOrBlank().not()) textCells++
            }
            textCells >= 2
        } ?: 0
        val header = sheet.getRow(headerRowIndex) ?: return ParsedLayout(nameIndex, 1, linkedMapOf())

        val columns = LinkedHashMap<String, Pair<Int, Int?>>()
        for (col in 0 until header.lastCellNum) {
            if (col == nameIndex) continue
            val headerText = header.getCell(col)?.toString()?.trim().orEmpty()
            if (headerText.isNotBlank()) columns[headerText] = col to null
        }
        return ParsedLayout(nameIndex, headerRowIndex + 1, columns)
    }

    private fun detectNameIndex(sheet: org.apache.poi.ss.usermodel.Sheet): Int {
        for (r in 0..5.coerceAtMost(sheet.lastRowNum)) {
            val row = sheet.getRow(r) ?: continue
            for (c in 0 until row.lastCellNum) {
                val value = row.getCell(c)?.toString()?.trim()?.lowercase().orEmpty()
                if (value == "student" || value == "name") return c
            }
        }
        return 1
    }

    private fun getNumericCellValue(cell: org.apache.poi.ss.usermodel.Cell?): Double? {
        if (cell == null) return null
        return when (cell.cellType) {
            CellType.NUMERIC -> cell.numericCellValue
            CellType.STRING -> cell.stringCellValue.toDoubleOrNull()
            else -> null
        }
    }

    private fun parseCsvLine(line: String): List<String> {
        val result = mutableListOf<String>()
        val sb = StringBuilder()
        var inQuotes = false
        var i = 0
        while (i < line.length) {
            val ch = line[i]
            when {
                ch == '"' -> {
                    if (inQuotes && i + 1 < line.length && line[i + 1] == '"') {
                        sb.append('"')
                        i++
                    } else {
                        inQuotes = !inQuotes
                    }
                }
                ch == ',' && !inQuotes -> {
                    result.add(sb.toString())
                    sb.setLength(0)
                }
                else -> sb.append(ch)
            }
            i++
        }
        result.add(sb.toString())
        return result
    }

    private fun buildCsvLayout(rows: List<List<String>>): ParsedLayout {
        val nameIndex = detectNameIndexFromCsv(rows)
        val marker = detectCoursesMarkerFromCsv(rows)

        if (marker != null) {
            val (coursesRowIndex, coursesStartCol) = marker
            val courseNameRow = rows.getOrNull(coursesRowIndex + 1).orEmpty()
            val subHeaderRow = rows.getOrNull(coursesRowIndex + 2).orEmpty()
            val maxCol = maxOf(courseNameRow.size, subHeaderRow.size)
            val hasCaExam = (coursesStartCol until maxCol).any { col ->
                val value = subHeaderRow.getOrNull(col)?.trim()?.lowercase().orEmpty()
                value == "ca" || value == "exam"
            }

            val columns = linkedMapOf<String, Pair<Int, Int?>>()
            var currentCourse = ""
            for (col in coursesStartCol until maxCol) {
                val candidate = courseNameRow.getOrNull(col)?.trim().orEmpty()
                if (candidate.isNotBlank() && candidate.lowercase() !in setOf("courses", "ca", "exam")) {
                    currentCourse = candidate
                    if (!hasCaExam && col != nameIndex) {
                        columns[currentCourse] = col to null
                    }
                }
                if (!hasCaExam || currentCourse.isBlank()) continue

                when (subHeaderRow.getOrNull(col)?.trim()?.lowercase()) {
                    "ca" -> {
                        val exam = columns[currentCourse]?.second
                        columns[currentCourse] = col to exam
                    }
                    "exam" -> {
                        val ca = columns[currentCourse]?.first ?: -1
                        columns[currentCourse] = ca to col
                    }
                }
            }

            if (hasCaExam) {
                columns.entries.removeIf { (_, cols) ->
                    val (caCol, examCol) = cols
                    caCol < 0 || examCol == null || examCol < 0
                }
            }

            if (columns.isNotEmpty()) {
                val dataStart = if (hasCaExam) coursesRowIndex + 3 else coursesRowIndex + 2
                return ParsedLayout(nameIndex, dataStart, LinkedHashMap(columns))
            }
        }

        val header = rows.firstOrNull().orEmpty()
        val columns = linkedMapOf<String, Pair<Int, Int?>>()
        for ((col, headerTextRaw) in header.withIndex()) {
            if (col == nameIndex) continue
            val headerText = headerTextRaw.trim()
            if (headerText.isNotBlank()) columns[headerText] = col to null
        }
        return ParsedLayout(nameIndex, 1, LinkedHashMap(columns))
    }

    private fun detectNameIndexFromCsv(rows: List<List<String>>): Int {
        for (row in rows.take(3)) {
            for ((col, value) in row.withIndex()) {
                if (value.trim().lowercase() in setOf("student", "name")) return col
            }
        }
        return if (rows.any { it.size > 1 }) 1 else 0
    }

    private fun detectCoursesMarkerFromSheet(sheet: org.apache.poi.ss.usermodel.Sheet): Pair<Int, Int>? {
        for (r in 0..5.coerceAtMost(sheet.lastRowNum)) {
            val row = sheet.getRow(r) ?: continue
            for (c in 0 until row.lastCellNum) {
                val value = row.getCell(c)?.toString()?.trim()?.lowercase().orEmpty()
                if (value == "courses") return r to c
            }
        }
        return null
    }

    private fun detectCoursesMarkerFromCsv(rows: List<List<String>>): Pair<Int, Int>? {
        for ((r, row) in rows.take(6).withIndex()) {
            for ((c, value) in row.withIndex()) {
                if (value.trim().lowercase() == "courses") return r to c
            }
        }
        return null
    }
}

// ========================== Data Class for Results ==========================
data class StudentResult(val student: String, val courseTotals: Map<String, Double>)

// ========================== Desktop GUI ==========================
class GradeCalculatorGUI : JFrame("Grade Calculator") {
    private val calculator = GradeCalculator()
    private val tabbedPane = JTabbedPane()
    private val originalPanel = JPanel(BorderLayout())
    private val gradesPanel = JPanel(BorderLayout())
    private val originalTextArea = JTextArea().apply {
        font = Font("Monospaced", Font.PLAIN, 12)
        isEditable = false
    }
    private val gradesTextArea = JTextArea().apply {
        font = Font("Monospaced", Font.PLAIN, 12)
        isEditable = false
    }
    private val statusBar = JLabel("Ready")

    init {
        defaultCloseOperation = EXIT_ON_CLOSE
        setLayout(BorderLayout())

        // Button panel
        val buttonPanel = JPanel(FlowLayout())
        val manualBtn = JButton("Manual Entry")
        val fileBtn = JButton("Open File...")
        val urlBtn = JButton("Download from URL")
        buttonPanel.add(manualBtn)
        buttonPanel.add(fileBtn)
        buttonPanel.add(urlBtn)

        // Tabbed pane with original and grades sheets
        originalPanel.add(JScrollPane(originalTextArea), BorderLayout.CENTER)
        gradesPanel.add(JScrollPane(gradesTextArea), BorderLayout.CENTER)
        tabbedPane.addTab("Original Scores", originalPanel)
        tabbedPane.addTab("Grades", gradesPanel)

        add(buttonPanel, BorderLayout.NORTH)
        add(tabbedPane, BorderLayout.CENTER)
        add(statusBar, BorderLayout.SOUTH)

        // Event handlers
        manualBtn.addActionListener { manualEntry() }
        fileBtn.addActionListener { chooseFile() }
        urlBtn.addActionListener { downloadFromUrl() }

        setSize(800, 600)
        setLocationRelativeTo(null)
    }

    private fun manualEntry() {
        val name = JOptionPane.showInputDialog(this, "Student name:")
        if (name.isNullOrBlank()) return
        val countStr = JOptionPane.showInputDialog(this, "How many scores?")
        val count = countStr?.toIntOrNull() ?: return
        val scores = mutableListOf<Double>()
        for (i in 1..count) {
            val scoreStr = JOptionPane.showInputDialog(this, "Enter score $i:")
            val score = scoreStr?.toDoubleOrNull() ?: return
            scores.add(score)
        }

        val (avg, grade) = calculator.processScores(scores) { list -> list.sum() / list.size }
        val result = "$name -> Average: %.2f, Grade: $grade".format(avg)

        val bonus = 2.0
        val avgWithBonus = calculator.processScoresWithOperation(
            scores,
            operation = { list -> list.map { it + bonus } },
            aggregator = { list -> list.sum() / list.size }
        )
        val bonusResult = "With +$bonus bonus: Average = %.2f, Grade = %s".format(avgWithBonus, calculator.calculateGrade(avgWithBonus))

        // Display in a dialog or in the original tab? For simplicity, show in a message.
        JOptionPane.showMessageDialog(this, "$result\n$bonusResult", "Manual Entry Result", JOptionPane.INFORMATION_MESSAGE)
    }

    private fun chooseFile() {
        val chooser = JFileChooser()
        if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            val file = chooser.selectedFile
            statusBar.text = "Processing ${file.name}..."
            try {
                val results = calculator.processFile(file)
                if (results is List<*>) {
                    // After processing, load the saved Excel file and display it
                    val savedFile = File("student_grades.xlsx")
                    if (savedFile.exists()) {
                        displayExcelFile(savedFile)
                        statusBar.text = "Processed ${file.name} – saved as ${savedFile.absolutePath}"
                    } else {
                        statusBar.text = "Processing completed, but output file not found."
                    }
                } else {
                    statusBar.text = "No results from file."
                }
            } catch (ex: Exception) {
                statusBar.text = "Error: ${ex.message}"
                JOptionPane.showMessageDialog(this, "Error processing file:\n${ex.message}", "Error", JOptionPane.ERROR_MESSAGE)
            }
        }
    }

    private fun downloadFromUrl() {
        val url = JOptionPane.showInputDialog(this, "Enter URL (Excel or CSV):")
        if (url.isNullOrBlank()) return
        statusBar.text = "Downloading..."
        try {
            val file = calculator.download(url)
            statusBar.text = "Downloaded to ${file.absolutePath}. Processing..."
            val results = calculator.processFile(file)
            if (results is List<*>) {
                val savedFile = File("student_grades.xlsx")
                if (savedFile.exists()) {
                    displayExcelFile(savedFile)
                    statusBar.text = "Processed downloaded file – saved as ${savedFile.absolutePath}"
                } else {
                    statusBar.text = "Processing completed, but output file not found."
                }
                file.deleteOnExit()
            }
        } catch (ex: Exception) {
            statusBar.text = "Download failed: ${ex.message}"
            JOptionPane.showMessageDialog(this, "Download failed:\n${ex.message}", "Error", JOptionPane.ERROR_MESSAGE)
        }
    }

    private fun displayExcelFile(file: File) {
        try {
            FileInputStream(file).use { fis ->
                XSSFWorkbook(fis).use { workbook ->
                    // Display original sheet (first sheet or "Scores")
                    val scoresSheet = workbook.getSheet("Scores") ?: workbook.getSheetAt(0)
                    originalTextArea.text = sheetToText(scoresSheet)

                    // Display grades sheet
                    val gradesSheet = workbook.getSheet("Grades")
                    gradesTextArea.text = if (gradesSheet != null) sheetToText(gradesSheet) else "No Grades sheet found."
                }
            }
        } catch (ex: Exception) {
            originalTextArea.text = "Error loading Excel file: ${ex.message}"
            gradesTextArea.text = ""
        }
    }

    private fun sheetToText(sheet: org.apache.poi.ss.usermodel.Sheet): String {
        if (sheet.lastRowNum < 0) return "(empty sheet)"
        val colWidths = mutableMapOf<Int, Int>()
        for (r in 0..sheet.lastRowNum) {
            val row = sheet.getRow(r) ?: continue
            for (c in 0 until row.lastCellNum) {
                val cell = row.getCell(c)
                val value = when (cell?.cellType) {
                    CellType.NUMERIC -> String.format("%.2f", cell.numericCellValue)
                    CellType.STRING -> cell.stringCellValue
                    CellType.BOOLEAN -> cell.booleanCellValue.toString()
                    else -> ""
                }
                colWidths[c] = maxOf(colWidths[c] ?: 0, value.length)
            }
        }
        val sb = StringBuilder()
        fun separator() {
            sb.append("+")
            for (c in 0 until colWidths.size) {
                sb.append("-".repeat(colWidths[c]!! + 2)).append("+")
            }
            sb.append("\n")
        }
        separator()
        for (r in 0..sheet.lastRowNum) {
            val row = sheet.getRow(r) ?: continue
            sb.append("|")
            for (c in 0 until colWidths.size) {
                val cell = row.getCell(c)
                val value = when (cell?.cellType) {
                    CellType.NUMERIC -> String.format("%.2f", cell.numericCellValue)
                    CellType.STRING -> cell.stringCellValue
                    CellType.BOOLEAN -> cell.booleanCellValue.toString()
                    else -> ""
                }
                sb.append(" %-${colWidths[c]!!}s |".format(value))
            }
            sb.append("\n")
            separator()
        }
        return sb.toString()
    }
}

// ========================== Main Entry Point ==========================
fun main(args: Array<String>) {
    // Always launch GUI if not headless, otherwise fallback to console.
    if (!GraphicsEnvironment.isHeadless()) {
        SwingUtilities.invokeLater {
            GradeCalculatorGUI().isVisible = true
        }
    } else {
        println("Headless environment – falling back to console mode.")
        val calculator = GradeCalculator()
        ConsoleMenu(calculator).run()
    }
}

// Keep the ConsoleMenu class from before (unchanged) – it's used only in headless mode.
class ConsoleMenu(private val calculator: GradeCalculator) {
    private val scanner = Scanner(System.`in`)

    fun run() {
        while (true) {
            println("\n=== Student Grade Calculator (Console) ===")
            println("1. Enter scores manually")
            println("2. Process a local file (.xlsx, .csv, .pdf, .html)")
            println("3. Exit")
            println("4. Download and process from a URL (Excel/CSV)")
            print("Choose an option: ")

            when (scanner.nextLine().trim()) {
                "1" -> manualEntry()
                "2" -> processLocalFile()
                "3" -> {
                    println("Goodbye!")
                    return
                }
                "4" -> downloadFromUrl()
                else -> println("Invalid option.")
            }
        }
    }

    private fun manualEntry() {
        print("Student name: ")
        val name = scanner.nextLine().trim()
        if (name.isBlank()) return
        print("How many scores? ")
        val count = scanner.nextLine().trim().toIntOrNull() ?: return
        val scores = mutableListOf<Double>()
        for (i in 1..count) {
            print("Enter score $i: ")
            val score = scanner.nextLine().trim().toDoubleOrNull() ?: return
            scores.add(score)
        }

        val (avg, grade) = calculator.processScores(scores) { list -> list.sum() / list.size }
        println("$name -> Average: %.2f, Grade: $grade".format(avg))

        val bonus = 2.0
        val avgWithBonus = calculator.processScoresWithOperation(
            scores,
            operation = { list -> list.map { it + bonus } },
            aggregator = { list -> list.sum() / list.size }
        )
        println("With +$bonus bonus: Average = %.2f, Grade = %s".format(avgWithBonus, calculator.calculateGrade(avgWithBonus)))
    }

    private fun processLocalFile() {
        print("Enter file path: ")
        val path = scanner.nextLine().trim()
        val file = File(path)
        if (!file.exists()) {
            println("File not found.")
            return
        }
        val results = calculator.processFile(file)
        if (results is List<*>) {
            // Tables are already printed to console by processExcel/processCsv
            val pdf = calculator.generateOutput(results, OutputFormat.PDF)
            println("Generated PDF: ${pdf.absolutePath}")
            val html = calculator.generateOutput(results, OutputFormat.HTML)
            println("Generated HTML: ${html.absolutePath}")
        }
    }

    private fun downloadFromUrl() {
        print("Enter URL (Excel or CSV): ")
        val url = scanner.nextLine().trim()
        try {
            val file = calculator.download(url)
            val results = calculator.processFile(file)
            if (results is List<*>) {
                file.deleteOnExit()
                val pdf = calculator.generateOutput(results, OutputFormat.PDF)
                println("Generated PDF: ${pdf.absolutePath}")
                val html = calculator.generateOutput(results, OutputFormat.HTML)
                println("Generated HTML: ${html.absolutePath}")
            }
        } catch (e: Exception) {
            println("Download failed: ${e.message}")
        }
    }
}
