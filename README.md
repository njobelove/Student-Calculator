# Student Grade Calculator (Kotlin)

This repository now contains:

- a Kotlin console app (Excel + console input),
- an Android mobile app version.

## Project Structure

- `src/main/kotlin/org/example/Main.kt`
- `run-console.bat` (safe Windows launcher for console app)
- `android-mobile/` (Android Studio mobile app project)

## Console App Features

- Manual console grading from entered scores.
- Excel/CSV mode:
  - reads local or online `.xlsx` or `.csv`,
  - creates `Grades` output,
  - writes `student_grades.xlsx` or `student_grades.csv`,
  - prints both input and `Grades` to console.

CSV notes (structure-tolerant):
- Student name is always column 2.
- Only columns under a header labeled `Courses` are used as scores.
- Adding/removing rows or columns will not break grading as long as the `Courses` header exists.

## GradeCalculator Class (Console)

- Property:
  - `passMark: Double` (default `60.0`)
- Structure:
  - `calculateGrade(average: Double): String`
  - `calculateAverage(scores: List<Double>): Double`
  - `createSampleScoresFile(path: String)`
  - `processExcel(inputPath: String, outputPath: String)`
  - `processCsv(inputPath: String, outputPath: String)`
  - `printSheet(workbook: XSSFWorkbook, sheetName: String)`
- Main procedure:
  - select mode `1` (console), `2` (local Excel/CSV), or `4` (online Excel/CSV),
  - execute calculation,
  - print result to console.

## Run Console App in IntelliJ IDEA

1. Open this root project in IntelliJ IDEA.
2. Wait for Gradle sync.
3. Open `src/main/kotlin/org/example/Main.kt`.
4. Run `main()`.
5. Choose:
   - `2` for local Excel/CSV file
   - `4` for online Excel/CSV URL

## Run Console App from Terminal (Windows)

From the project root, use:

```bat
run-console.bat
```

Alternative:

```bat
gradlew.bat --console=plain run
```

If IntelliJ terminal says command is inaccessible:

1. Make sure terminal starts in project root:
   - `C:\Users\user\Desktop\StudentCalculator\StudentCalculator`
2. Use `gradlew.bat` (not `gradlew`).
3. If using PowerShell terminal, run:
   - `cmd /c gradlew.bat run`
4. If still blocked, use IntelliJ Run button on `main()`.

## Android Mobile App

Android app location:

- `android-mobile/`

Main files:

- `android-mobile/app/src/main/java/org/example/studentcalculatormobile/MainActivity.kt`
- `android-mobile/app/src/main/java/org/example/studentcalculatormobile/GradeCalculator.kt`
- `android-mobile/app/src/main/res/layout/activity_main.xml`

Features:

- Paste online downloadable `.xlsx` URL.
- Optionally set input sheet name (default `Scores`).
- Tap `Process Online Excel`.
- App downloads input file, creates `Grades` sheet, saves output workbook.
- App displays preview of both `Scores` and `Grades` sheets.

## Run Android App (Android Studio)

1. Open Android Studio.
2. Choose `Open` and select `android-mobile` folder.
3. Let Gradle sync complete.
4. Create/start an emulator (or connect Android device with USB debugging).
5. Click `Run`.

## Important for Excel Online Links

Use a direct downloadable `.xlsx` link. If your current share link opens in browser only, first download/export the workbook link from OneDrive/Excel Online, then paste that downloadable URL into the app.
