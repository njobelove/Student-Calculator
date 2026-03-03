package org.example.studentcalculatormobile

import android.os.Bundle
import android.os.Environment
import android.widget.EditText
import android.widget.TextView
import androidx.appcompat.app.AppCompatActivity
import androidx.lifecycle.lifecycleScope
import com.google.android.material.button.MaterialButton
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.net.URI

class MainActivity : AppCompatActivity() {
    private val calculator = GradeCalculator()

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        val excelUrlInput = findViewById<EditText>(R.id.excelUrlInput)
        val sheetNameInput = findViewById<EditText>(R.id.sheetNameInput)
        val processExcelButton = findViewById<MaterialButton>(R.id.processExcelButton)
        val resultText = findViewById<TextView>(R.id.resultText)

        processExcelButton.setOnClickListener {
            val url = excelUrlInput.text.toString().trim()
            val sheetName = sheetNameInput.text.toString().trim().ifBlank { "Scores" }

            if (url.isBlank()) {
                resultText.text = "Enter a valid downloadable .xlsx URL."
                return@setOnClickListener
            }

            resultText.text = "Processing online Excel..."

            lifecycleScope.launch {
                val message = withContext(Dispatchers.IO) {
                    processOnlineExcel(url, sheetName)
                }
                resultText.text = message
            }
        }
    }

    private fun processOnlineExcel(url: String, scoreSheetName: String): String {
        return try {
            val downloadsDir = getExternalFilesDir(Environment.DIRECTORY_DOWNLOADS) ?: filesDir
            val inputFile = File(downloadsDir, "scores_online.xlsx")
            val outputFile = File(downloadsDir, "student_grades_online.xlsx")

            URI(url).toURL().openStream().use { input ->
                FileOutputStream(inputFile).use { output ->
                    input.copyTo(output)
                }
            }

            val preview = processWorkbook(inputFile, outputFile, scoreSheetName)
            "Done.\nInput: ${inputFile.absolutePath}\nOutput: ${outputFile.absolutePath}\n\n$preview"
        } catch (ex: Exception) {
            "Failed: ${ex.message}\nTip: use a direct downloadable .xlsx link."
        }
    }

    private fun processWorkbook(inputFile: File, outputFile: File, scoreSheetName: String): String {
        FileInputStream(inputFile).use { fis ->
            XSSFWorkbook(fis).use { workbook ->
                val scoresSheet = workbook.getSheet(scoreSheetName) ?: workbook.getSheetAt(0)

                val existingGrades = workbook.getSheet("Grades")
                if (existingGrades != null) {
                    workbook.removeSheetAt(workbook.getSheetIndex(existingGrades))
                }
                val gradesSheet = workbook.createSheet("Grades")

                val header = gradesSheet.createRow(0)
                header.createCell(0).setCellValue("Student")
                header.createCell(1).setCellValue("Average")
                header.createCell(2).setCellValue("Grade")

                var outRow = 1
                for (i in 1..scoresSheet.lastRowNum) {
                    val row = scoresSheet.getRow(i) ?: continue
                    val student = row.getCell(0)?.toString()?.trim().orEmpty()
                    if (student.isBlank()) continue

                    val scores = mutableListOf<Double>()
                    for (col in 1 until row.lastCellNum) {
                        val cell = row.getCell(col) ?: continue
                        when (cell.cellType) {
                            CellType.NUMERIC -> scores.add(cell.numericCellValue)
                            CellType.STRING -> cell.stringCellValue.toDoubleOrNull()?.let { scores.add(it) }
                            else -> {}
                        }
                    }

                    val avg = calculator.calculateAverage(scores)
                    val grade = calculator.calculateGrade(avg)

                    val gRow = gradesSheet.createRow(outRow++)
                    gRow.createCell(0).setCellValue(student)
                    gRow.createCell(1).setCellValue(String.format("%.2f", avg))
                    gRow.createCell(2).setCellValue(grade)
                }

                FileOutputStream(outputFile).use { fos -> workbook.write(fos) }
                val scoresPreview = sheetPreview(scoresSheet)
                val gradesPreview = sheetPreview(gradesSheet)
                return "Scores sheet:\n$scoresPreview\n\nGrades sheet:\n$gradesPreview"
            }
        }
    }

    private fun sheetPreview(sheet: org.apache.poi.ss.usermodel.Sheet): String {
        val lines = mutableListOf<String>()
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
            lines.add(values.joinToString(" | "))
        }
        return lines.joinToString("\n")
    }
}
