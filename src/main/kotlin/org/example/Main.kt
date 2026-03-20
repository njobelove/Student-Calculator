package org.example

import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.awt.*
import java.awt.event.*
import java.awt.geom.*
import java.io.*
import java.net.URI
import java.util.Scanner
import javax.swing.*
import javax.swing.border.*
import javax.swing.table.*

// ══════════════════════════════════════════════════════════════
//  DESIGN TOKENS
// ══════════════════════════════════════════════════════════════

object Theme {
    val BG_DEEP        = Color(0x0D1117)
    val BG_PANEL       = Color(0x161B22)
    val BG_CARD        = Color(0x1C2333)
    val BG_INPUT       = Color(0x0D1117)
    val BORDER         = Color(0x30363D)
    val ACCENT         = Color(0x2DD4BF)
    val ACCENT_DIM     = Color(0x1A7A70)
    val AMBER          = Color(0xF59E0B)
    val RED            = Color(0xEF4444)
    val GREEN          = Color(0x22C55E)
    val TEXT_PRIMARY   = Color(0xF0F6FC)
    val TEXT_SECONDARY = Color(0x8B949E)
    val TEXT_MUTED     = Color(0x484F58)

    val FONT_TITLE   = Font("SansSerif", Font.BOLD, 22)
    val FONT_HEADING = Font("SansSerif", Font.BOLD, 14)
    val FONT_BODY    = Font("SansSerif", Font.PLAIN, 13)
    val FONT_SMALL   = Font("SansSerif", Font.PLAIN, 11)
    val FONT_MONO    = Font("Monospaced", Font.PLAIN, 12)
    val FONT_BADGE   = Font("SansSerif", Font.BOLD, 11)

    fun gradeColor(grade: String) = when (grade) {
        "A"  -> Color(0x22C55E)
        "B"  -> Color(0x3B82F6)
        "C"  -> Color(0xF59E0B)
        "D"  -> Color(0xF97316)
        else -> Color(0xEF4444)
    }
}

// ══════════════════════════════════════════════════════════════
//  CUSTOM UI COMPONENTS
// ══════════════════════════════════════════════════════════════

/** Rounded card panel with an optional gradient header strip. */
class CardPanel(
    private val headerText: String? = null,
    private val headerIcon: String? = null,
    radius: Int = 12
) : JPanel(BorderLayout()) {

    private val arc = radius

    init {
        isOpaque = false
        border = EmptyBorder(if (headerText != null) 44 else 16, 16, 16, 16)
    }

    override fun paintComponent(g: Graphics) {
        val g2 = g as Graphics2D
        g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON)

        val shape = RoundRectangle2D.Float(0f, 0f, width.toFloat(), height.toFloat(), arc.toFloat(), arc.toFloat())
        g2.color = Theme.BG_CARD
        g2.fill(shape)
        g2.color = Theme.BORDER
        g2.draw(shape)

        if (headerText != null) {
            g2.paint = GradientPaint(0f, 0f, Theme.ACCENT_DIM, width.toFloat(), 0f, Color(0x0D1117))
            g2.fill(RoundRectangle2D.Float(0f, 0f, width.toFloat(), 36f, arc.toFloat(), arc.toFloat()))
            g2.color = Theme.BG_CARD
            g2.fillRect(0, 20, width, 16)
            g2.font = Theme.FONT_HEADING
            g2.color = Theme.TEXT_PRIMARY
            val prefix = if (headerIcon != null) "$headerIcon  " else ""
            g2.drawString("$prefix$headerText", 16, 24)
        }

        super.paintComponent(g)
    }
}

/** Filled pill-shaped button with hover brightening. */
class StyledButton(
    text: String,
    private val bg: Color = Theme.ACCENT,
    private val fg: Color = Theme.BG_DEEP
) : JButton(text) {

    private var hovered = false

    init {
        isContentAreaFilled = false
        isFocusPainted = false
        isBorderPainted = false
        font = Font("SansSerif", Font.BOLD, 13)
        foreground = fg
        cursor = Cursor.getPredefinedCursor(Cursor.HAND_CURSOR)
        preferredSize = Dimension(preferredSize.width + 32, 38)
        addMouseListener(object : MouseAdapter() {
            override fun mouseEntered(e: MouseEvent) { hovered = true;  repaint() }
            override fun mouseExited(e: MouseEvent)  { hovered = false; repaint() }
        })
    }

    override fun paintComponent(g: Graphics) {
        val g2 = g as Graphics2D
        g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON)
        g2.color = if (hovered) bg.brighter() else bg
        g2.fill(RoundRectangle2D.Float(0f, 0f, width.toFloat(), height.toFloat(), 20f, 20f))
        g2.font = font
        g2.color = fg
        val fm = g2.fontMetrics
        g2.drawString(text, (width - fm.stringWidth(text)) / 2, (height + fm.ascent - fm.descent) / 2)
    }
}

/** Outlined ghost button with hover fill. */
class GhostButton(text: String, private val accent: Color = Theme.ACCENT) : JButton(text) {

    private var hovered = false

    init {
        isContentAreaFilled = false
        isFocusPainted = false
        isBorderPainted = false
        font = Font("SansSerif", Font.BOLD, 13)
        foreground = accent
        cursor = Cursor.getPredefinedCursor(Cursor.HAND_CURSOR)
        preferredSize = Dimension(preferredSize.width + 32, 38)
        addMouseListener(object : MouseAdapter() {
            override fun mouseEntered(e: MouseEvent) { hovered = true;  repaint() }
            override fun mouseExited(e: MouseEvent)  { hovered = false; repaint() }
        })
    }

    override fun paintComponent(g: Graphics) {
        val g2 = g as Graphics2D
        g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON)
        if (hovered) {
            g2.color = Color(accent.red, accent.green, accent.blue, 30)
            g2.fill(RoundRectangle2D.Float(0f, 0f, width.toFloat(), height.toFloat(), 20f, 20f))
        }
        g2.color = accent
        g2.stroke = BasicStroke(1.5f)
        g2.draw(RoundRectangle2D.Float(1f, 1f, width - 2f, height - 2f, 20f, 20f))
        g2.font = font
        g2.color = if (hovered) accent.brighter() else accent
        val fm = g2.fontMetrics
        g2.drawString(text, (width - fm.stringWidth(text)) / 2, (height + fm.ascent - fm.descent) / 2)
    }
}

/** Dark text field with a rounded focus border. */
class StyledTextField(cols: Int = 20) : JTextField(cols) {
    init {
        background = Theme.BG_INPUT
        foreground = Theme.TEXT_PRIMARY
        caretColor = Theme.ACCENT
        font = Theme.FONT_BODY
        isOpaque = true
        border = CompoundBorder(
            object : AbstractBorder() {
                override fun paintBorder(c: Component, g: Graphics, x: Int, y: Int, w: Int, h: Int) {
                    val g2 = g as Graphics2D
                    g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON)
                    g2.color = if (c.hasFocus()) Theme.ACCENT else Theme.BORDER
                    g2.drawRoundRect(x, y, w - 1, h - 1, 8, 8)
                }
                override fun getBorderInsets(c: Component) = Insets(1, 1, 1, 1)
            },
            EmptyBorder(6, 10, 6, 10)
        )
    }
}

/** Colour-coded grade badge label. */
class GradeBadge(grade: String) : JLabel(grade, JLabel.CENTER) {
    init {
        val c = Theme.gradeColor(grade)
        foreground = c
        font = Theme.FONT_BADGE
        isOpaque = false
        border = CompoundBorder(
            object : AbstractBorder() {
                override fun paintBorder(comp: Component, g: Graphics, x: Int, y: Int, w: Int, h: Int) {
                    val g2 = g as Graphics2D
                    g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON)
                    g2.color = c
                    g2.drawRoundRect(x, y, w - 1, h - 1, 10, 10)
                }
                override fun getBorderInsets(c: Component) = Insets(1, 1, 1, 1)
            },
            EmptyBorder(2, 8, 2, 8)
        )
    }
}

/** Bottom status bar with a coloured dot indicator. */
class StatusBar : JPanel(BorderLayout()) {

    enum class StatusType { INFO, OK, ERROR }

    private val dot      = JLabel("●")
    private val msgLabel = JLabel("  Ready")

    init {
        isOpaque = false
        preferredSize = Dimension(0, 28)
        dot.foreground = Theme.GREEN
        dot.font = Theme.FONT_SMALL
        dot.border = EmptyBorder(0, 8, 0, 4)
        msgLabel.font = Theme.FONT_SMALL
        msgLabel.foreground = Theme.TEXT_SECONDARY
        add(dot, BorderLayout.WEST)
        add(msgLabel, BorderLayout.CENTER)
    }

    override fun paintComponent(g: Graphics) {
        g.color = Theme.BG_PANEL
        g.fillRect(0, 0, width, height)
        g.color = Theme.BORDER
        (g as Graphics2D).drawLine(0, 0, width, 0)
        super.paintComponent(g)
    }

    fun setText(msg: String, type: StatusType = StatusType.INFO) {
        dot.foreground = when (type) {
            StatusType.OK    -> Theme.GREEN
            StatusType.ERROR -> Theme.RED
            else             -> Theme.ACCENT
        }
        msgLabel.text = "  $msg"
    }
}

/** Dark-themed JTable with alternating rows and coloured grade cells. */
class DarkTable(model: TableModel) : JTable(model) {
    init {
        background = Theme.BG_CARD
        foreground = Theme.TEXT_PRIMARY
        gridColor = Theme.BORDER
        font = Theme.FONT_MONO
        rowHeight = 28
        showHorizontalLines = true
        showVerticalLines = false
        selectionBackground = Color(Theme.ACCENT.red, Theme.ACCENT.green, Theme.ACCENT.blue, 50)
        selectionForeground = Theme.TEXT_PRIMARY
        tableHeader.background = Theme.BG_PANEL
        tableHeader.foreground = Theme.TEXT_SECONDARY
        tableHeader.font = Font("SansSerif", Font.BOLD, 12)
        (tableHeader as JTableHeader).border = MatteBorder(0, 0, 1, 0, Theme.BORDER)

        setDefaultRenderer(Any::class.java, object : DefaultTableCellRenderer() {
            override fun getTableCellRendererComponent(
                t: JTable, value: Any?, sel: Boolean, focus: Boolean, row: Int, col: Int
            ): Component {
                val c = super.getTableCellRendererComponent(t, value, sel, focus, row, col) as JLabel
                val str = value?.toString() ?: ""
                val isGrade = str in listOf("A", "B", "C", "D", "F")
                c.background = if (sel) selectionBackground else if (row % 2 == 0) Theme.BG_CARD else Theme.BG_PANEL
                c.foreground = when {
                    sel     -> Theme.TEXT_PRIMARY
                    isGrade -> Theme.gradeColor(str)
                    else    -> Theme.TEXT_PRIMARY
                }
                c.font = if (isGrade) Font("SansSerif", Font.BOLD, 13) else Theme.FONT_MONO
                c.border = EmptyBorder(0, 10, 0, 10)
                return c
            }
        })
    }
}

/** Styled scroll pane matching the dark theme. */
fun darkScroll(view: Component): JScrollPane {
    val sp = JScrollPane(view)
    sp.background = Theme.BG_CARD
    sp.viewport.background = Theme.BG_CARD
    sp.border = MatteBorder(1, 1, 1, 1, Theme.BORDER)
    return sp
}

// ══════════════════════════════════════════════════════════════
//  SIDEBAR BUTTON
// ══════════════════════════════════════════════════════════════

/** A sidebar navigation item that highlights when selected. */
class SidebarButton(val label: String, val icon: String) : JPanel() {

    var selected = false
    var onClick: (() -> Unit)? = null
    private var hovered = false

    init {
        isOpaque = false
        cursor = Cursor.getPredefinedCursor(Cursor.HAND_CURSOR)
        preferredSize = Dimension(200, 44)
        addMouseListener(object : MouseAdapter() {
            override fun mouseEntered(e: MouseEvent) { hovered = true;  repaint() }
            override fun mouseExited(e: MouseEvent)  { hovered = false; repaint() }
            override fun mouseClicked(e: MouseEvent) { onClick?.invoke() }
        })
    }

    override fun paintComponent(g: Graphics) {
        val g2 = g as Graphics2D
        g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON)

        when {
            selected -> {
                g2.color = Color(Theme.ACCENT.red, Theme.ACCENT.green, Theme.ACCENT.blue, 25)
                g2.fill(RoundRectangle2D.Float(4f, 2f, width - 8f, height - 4f, 10f, 10f))
                g2.color = Theme.ACCENT
                g2.stroke = BasicStroke(2f)
                g2.drawLine(0, 8, 0, height - 8)
            }
            hovered -> {
                g2.color = Color(255, 255, 255, 12)
                g2.fill(RoundRectangle2D.Float(4f, 2f, width - 8f, height - 4f, 10f, 10f))
            }
        }

        g2.font = Font("SansSerif", Font.PLAIN, 18)
        g2.color = if (selected) Theme.ACCENT else Theme.TEXT_SECONDARY
        g2.drawString(icon, 16, height / 2 + 7)

        g2.font = if (selected) Font("SansSerif", Font.BOLD, 13) else Theme.FONT_BODY
        g2.color = if (selected) Theme.TEXT_PRIMARY else Theme.TEXT_SECONDARY
        g2.drawString(label, 46, height / 2 + 5)
    }
}

// ══════════════════════════════════════════════════════════════
//  MAIN GUI
// ══════════════════════════════════════════════════════════════

class GradeCalculatorGUI : JFrame("Grade Calculator") {

    private val calculator       = GradeCalculator()
    private val statusBar        = StatusBar()
    private val contentArea      = JPanel(CardLayout())
    private var lastResults      = emptyList<StudentResult>()
    private val resultsTableModel = DefaultTableModel()
    private var resultsScoresTab: JPanel? = null
    private var lastUploadedFile: File?   = null

    private val sidebarBtns = listOf(
        SidebarButton("Dashboard",    "⊞"),
        SidebarButton("Manual Entry", "✎"),
        SidebarButton("Open File",    "📂"),
        SidebarButton("Download URL", "⬇"),
        SidebarButton("Results",      "📊")
    )

    init {
        defaultCloseOperation = EXIT_ON_CLOSE
        try { UIManager.setLookAndFeel(UIManager.getCrossPlatformLookAndFeelClassName()) } catch (_: Exception) {}
        contentPane.background = Theme.BG_DEEP
        layout = BorderLayout()
        add(buildSidebar(),   BorderLayout.WEST)
        add(buildMainArea(),  BorderLayout.CENTER)
        add(statusBar,        BorderLayout.SOUTH)
        selectPage(0)
        setSize(1040, 680)
        setLocationRelativeTo(null)
        minimumSize = Dimension(820, 520)
    }

    // ── Sidebar ───────────────────────────────────────────────

    private fun buildSidebar(): JPanel {
        val sidebar = object : JPanel() {
            override fun paintComponent(g: Graphics) {
                g.color = Theme.BG_PANEL
                g.fillRect(0, 0, width, height)
                g.color = Theme.BORDER
                (g as Graphics2D).drawLine(width - 1, 0, width - 1, height)
                super.paintComponent(g)
            }
        }
        sidebar.isOpaque = false
        sidebar.layout = BoxLayout(sidebar, BoxLayout.Y_AXIS)
        sidebar.preferredSize = Dimension(200, 0)

        // Gradient logo strip
        val logoPanel = object : JPanel(FlowLayout(FlowLayout.LEFT, 10, 0)) {
            override fun paintComponent(g: Graphics) {
                val g2 = g as Graphics2D
                g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON)
                g2.paint = GradientPaint(0f, 0f, Theme.ACCENT_DIM, width.toFloat(), 0f, Color(0x0D1117, true))
                g2.fillRect(0, 0, width, height)
                g2.color = Theme.BORDER
                g2.drawLine(0, height - 1, width, height - 1)
                super.paintComponent(g)
            }
        }
        logoPanel.isOpaque = false
        logoPanel.preferredSize = Dimension(200, 64)
        logoPanel.maximumSize  = Dimension(200, 64)
        logoPanel.border = EmptyBorder(0, 6, 0, 0)

        val iconLabel  = JLabel("🎓").also { it.font = Font("SansSerif", Font.PLAIN, 22) }
        val titleLabel = JLabel("GradeCalc").also {
            it.font = Font("SansSerif", Font.BOLD, 17)
            it.foreground = Theme.TEXT_PRIMARY
        }
        logoPanel.add(iconLabel)
        logoPanel.add(titleLabel)
        sidebar.add(logoPanel)

        // Navigation section
        val navSection = JPanel().also {
            it.isOpaque = false
            it.layout = BoxLayout(it, BoxLayout.Y_AXIS)
            it.border = EmptyBorder(16, 0, 0, 0)
        }
        val navLabel = JLabel("  NAVIGATION").also {
            it.font = Font("SansSerif", Font.BOLD, 10)
            it.foreground = Theme.TEXT_MUTED
            it.border = EmptyBorder(0, 16, 8, 0)
        }
        navSection.add(navLabel)
        sidebarBtns.forEachIndexed { i, btn ->
            btn.maximumSize = Dimension(200, 44)
            btn.onClick = { selectPage(i) }
            navSection.add(btn)
        }
        sidebar.add(navSection)
        sidebar.add(Box.createVerticalGlue())

        val versionLabel = JLabel("  v2.0  —  Grade Calculator").also {
            it.font = Theme.FONT_SMALL
            it.foreground = Theme.TEXT_MUTED
            it.border = EmptyBorder(0, 0, 12, 0)
            it.maximumSize = Dimension(200, 24)
        }
        sidebar.add(versionLabel)
        return sidebar
    }

    private fun selectPage(index: Int) {
        sidebarBtns.forEachIndexed { i, b -> b.selected = i == index; b.repaint() }
        (contentArea.layout as CardLayout).show(contentArea, index.toString())
    }

    // ── Main content area ─────────────────────────────────────

    private fun buildMainArea(): JPanel {
        contentArea.isOpaque = false
        contentArea.add(buildDashboardPage(), "0")
        contentArea.add(buildManualPage(),    "1")
        contentArea.add(buildFilePage(),      "2")
        contentArea.add(buildUrlPage(),       "3")
        contentArea.add(buildResultsPage(),   "4")
        return JPanel(BorderLayout()).also {
            it.isOpaque = false
            it.add(contentArea)
        }
    }

    // ── Pages ─────────────────────────────────────────────────

    private fun buildDashboardPage(): JPanel {
        val page    = darkPage()
        val wrapper = scrollWrapper(page)
        pageHeader(page, "Dashboard", "Overview & quick actions")

        // Stats row
        val statsRow = JPanel(GridLayout(1, 4, 16, 0)).also { it.isOpaque = false }
        statsRow.add(statCard("Total Students",  "—", Theme.ACCENT, "👤"))
        statsRow.add(statCard("Average Score",   "—", Theme.AMBER,  "📈"))
        statsRow.add(statCard("Pass Rate",       "—", Theme.GREEN,  "✅"))
        statsRow.add(statCard("Files Processed", "0", Theme.RED,    "📂"))
        page.add(padded(statsRow, 0, 0, 24, 0))

        // Quick actions
        val actCard  = CardPanel("Quick Actions", "⚡")
        val actPanel = JPanel(GridLayout(1, 3, 12, 0)).also { it.isOpaque = false }
        val manualBtn = StyledButton("✎  Manual Entry", Theme.ACCENT, Theme.BG_DEEP)
        val fileBtn   = StyledButton("📂  Open File",   Theme.AMBER,  Theme.BG_DEEP)
        val urlBtn    = StyledButton("⬇  Download URL", Theme.GREEN,  Theme.BG_DEEP)
        manualBtn.addActionListener { selectPage(1) }
        fileBtn.addActionListener   { selectPage(2) }
        urlBtn.addActionListener    { selectPage(3) }
        actPanel.add(manualBtn)
        actPanel.add(fileBtn)
        actPanel.add(urlBtn)
        actCard.add(actPanel)
        page.add(padded(actCard, 0, 0, 24, 0))

        // Grade scale legend
        val legendCard  = CardPanel("Grade Scale", "🎓")
        val legendPanel = JPanel(FlowLayout(FlowLayout.LEFT, 16, 8)).also { it.isOpaque = false }
        listOf("A" to "≥ 90", "B" to "≥ 80", "C" to "≥ 70", "D" to "≥ 60", "F" to "< 60").forEach { (grade, range) ->
            val item = JPanel(FlowLayout(FlowLayout.LEFT, 6, 0)).also { it.isOpaque = false }
            item.add(GradeBadge(grade))
            item.add(JLabel(range).also { it.foreground = Theme.TEXT_SECONDARY; it.font = Theme.FONT_SMALL })
            legendPanel.add(item)
        }
        legendCard.add(legendPanel)
        page.add(padded(legendCard, 0, 0, 0, 0))
        page.add(Box.createVerticalGlue())
        return wrapper
    }

    private fun buildManualPage(): JPanel {
        val page    = darkPage()
        val wrapper = scrollWrapper(page)
        pageHeader(page, "Manual Entry", "Enter student scores by hand")

        val formCard  = CardPanel("Student Details", "✎")
        val form      = JPanel(GridBagLayout()).also { it.isOpaque = false }
        val gbc       = GridBagConstraints().apply { insets = Insets(6, 6, 6, 6); fill = GridBagConstraints.HORIZONTAL }
        val nameField  = StyledTextField(20)
        val countField = StyledTextField(5)
        val bonusField = StyledTextField(5).also { it.text = "0" }
        val resultArea = JTextArea(6, 40).apply {
            background = Theme.BG_INPUT
            foreground = Theme.ACCENT
            caretColor = Theme.ACCENT
            font = Theme.FONT_MONO
            isEditable = false
            border = EmptyBorder(10, 12, 10, 12)
            lineWrap = true
        }

        gbc.gridx = 0; gbc.gridy = 0; gbc.weightx = 0.0; form.add(lbl("Student Name"), gbc)
        gbc.gridx = 1;                gbc.weightx = 1.0; form.add(nameField, gbc)
        gbc.gridx = 2;                gbc.weightx = 0.0; form.add(lbl("No. of Scores"), gbc)
        gbc.gridx = 3;                gbc.weightx = 0.5; form.add(countField, gbc)
        gbc.gridx = 0; gbc.gridy = 1; gbc.weightx = 0.0; form.add(lbl("Bonus Points"), gbc)
        gbc.gridx = 1;                gbc.weightx = 0.5; form.add(bonusField, gbc)

        val submitBtn = StyledButton("  Calculate Grades  ")
        val btnRow    = JPanel(FlowLayout(FlowLayout.LEFT)).also { it.isOpaque = false; it.add(submitBtn) }
        gbc.gridx = 0; gbc.gridy = 2; gbc.gridwidth = 4; form.add(btnRow, gbc)
        gbc.gridy = 3; form.add(JLabel("Result").also {
            it.font = Theme.FONT_HEADING; it.foreground = Theme.TEXT_SECONDARY; it.border = EmptyBorder(8, 0, 4, 0)
        }, gbc)
        gbc.gridy = 4; form.add(darkScroll(resultArea), gbc)

        formCard.add(form)
        page.add(padded(formCard, 0, 0, 0, 0))
        page.add(Box.createVerticalGlue())

        submitBtn.addActionListener {
            val name  = nameField.text.trim()
            val count = countField.text.trim().toIntOrNull()
            val bonus = bonusField.text.trim().toDoubleOrNull() ?: 0.0

            if (name.isBlank() || count == null || count <= 0) {
                resultArea.foreground = Theme.RED
                resultArea.text = "⚠  Please enter a valid name and score count."
                return@addActionListener
            }

            val scores = mutableListOf<Double>()
            for (i in 1..count) {
                val input = JOptionPane.showInputDialog(this, "Enter score $i of $count:", "Score Input", JOptionPane.PLAIN_MESSAGE)
                val s = input?.toDoubleOrNull()
                if (s == null) { resultArea.text = "Cancelled."; return@addActionListener }
                scores.add(s)
            }

            val (avg, grade) = calculator.processScores(scores) { list -> list.sum() / list.size }
            val avgBonus = calculator.processScoresWithOperation(
                scores,
                { l -> l.map { it + bonus } },
                { l -> l.sum() / l.size }
            )

            resultArea.foreground = Theme.gradeColor(grade)
            resultArea.text = buildString {
                appendLine("Student : $name")
                appendLine("Scores  : ${scores.joinToString(", ")}")
                appendLine("Average : ${"%.2f".format(avg)}")
                appendLine("Grade   : $grade")
                if (bonus > 0) {
                    appendLine("─────────────────────────")
                    appendLine("With +$bonus bonus  →  Avg: ${"%.2f".format(avgBonus)}, Grade: ${calculator.calculateGrade(avgBonus)}")
                }
            }

            lastResults = listOf(StudentResult(name, scores.mapIndexed { i, s -> "Score ${i + 1}" to s }.toMap()))
            try {
                calculator.generateOutput(lastResults, OutputFormat.EXCEL)
                statusBar.setText("Saved student_grades.xlsx", StatusBar.StatusType.OK)
            } catch (ex: Exception) {
                statusBar.setText("Save error: ${ex.message}", StatusBar.StatusType.ERROR)
            }
        }
        return wrapper
    }

    private fun buildFilePage(): JPanel {
        val page    = darkPage()
        val wrapper = scrollWrapper(page)
        pageHeader(page, "Open File", "Process .xlsx or .csv grade files")

        // ── Drop zone card ────────────────────────────────────
        val uploadCard = CardPanel("Upload File", "📂")
        val uploadInner = JPanel().also { it.isOpaque = false; it.layout = BoxLayout(it, BoxLayout.Y_AXIS) }

        val dropHint = object : JPanel() {
            override fun paintComponent(g: Graphics) {
                val g2 = g as Graphics2D
                g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON)
                g2.color = Color(Theme.ACCENT.red, Theme.ACCENT.green, Theme.ACCENT.blue, 15)
                g2.fill(RoundRectangle2D.Float(0f, 0f, width.toFloat(), height.toFloat(), 16f, 16f))
                g2.color = Theme.ACCENT_DIM
                g2.stroke = BasicStroke(1.5f, BasicStroke.CAP_BUTT, BasicStroke.JOIN_BEVEL, 0f, floatArrayOf(8f, 6f), 0f)
                g2.draw(RoundRectangle2D.Float(1f, 1f, width - 2f, height - 2f, 16f, 16f))
                super.paintComponent(g)
            }
        }
        dropHint.isOpaque = false
        dropHint.layout = BorderLayout()
        dropHint.preferredSize = Dimension(0, 100)
        dropHint.maximumSize  = Dimension(Int.MAX_VALUE, 100)
        dropHint.add(JLabel("Click 'Browse' to select a file   (.xlsx  /  .csv)", JLabel.CENTER).also {
            it.font = Theme.FONT_BODY
            it.foreground = Theme.TEXT_SECONDARY
        })

        val browseBtn = StyledButton("  📂  Browse File  ")
        val btnRow    = JPanel(FlowLayout(FlowLayout.LEFT)).also { it.isOpaque = false; it.add(browseBtn) }
        val statusLbl = JLabel("No file loaded").also { it.foreground = Theme.TEXT_MUTED; it.font = Theme.FONT_SMALL }

        uploadInner.add(dropHint)
        uploadInner.add(Box.createVerticalStrut(10))
        uploadInner.add(btnRow)
        uploadInner.add(Box.createVerticalStrut(4))
        uploadInner.add(statusLbl)
        uploadCard.add(uploadInner)
        page.add(padded(uploadCard, 0, 0, 16, 0))

        // ── Side-by-side score + grade tables ─────────────────
        val scoresTableModel = DefaultTableModel()
        val gradesTableModel = DefaultTableModel()

        val scoresCard = CardPanel("📄  Uploaded Scores", null)
        scoresCard.layout = BorderLayout()
        val scoresPlaceholder = placeholderLabel("Upload a file to see scores here")
        scoresCard.add(scoresPlaceholder, BorderLayout.CENTER)

        val gradesCard = CardPanel("✅  Generated Grades", null)
        gradesCard.layout = BorderLayout()
        val gradesPlaceholder = placeholderLabel("Grades will appear here after processing")
        gradesCard.add(gradesPlaceholder, BorderLayout.CENTER)

        val splitPanel = JPanel(GridLayout(1, 2, 16, 0)).also { it.isOpaque = false }
        splitPanel.add(scoresCard)
        splitPanel.add(gradesCard)

        val splitWrapper = JPanel(BorderLayout()).also {
            it.isOpaque = false
            it.preferredSize = Dimension(0, 380)
            it.add(splitPanel)
        }
        page.add(padded(splitWrapper, 0, 0, 0, 0))
        page.add(Box.createVerticalGlue())

        // ── Browse action ─────────────────────────────────────
        browseBtn.addActionListener {
            val chooser = JFileChooser()
            if (chooser.showOpenDialog(this) != JFileChooser.APPROVE_OPTION) return@addActionListener
            val file = chooser.selectedFile
            lastUploadedFile = file   // remember for Results page scores tab
            statusLbl.text      = "  Processing: ${file.name}…"
            statusLbl.foreground = Theme.AMBER
            statusBar.setText("Processing ${file.name}…")

            SwingUtilities.invokeLater {
                try {
                    // Load the scores from the uploaded file into the scores table BEFORE processing
                    if (file.extension.equals("xlsx", true)) {
                        populateTableModelFromExcel(file, scoresTableModel, sheetIndex = 0)
                    } else if (file.extension.equals("csv", true)) {
                        populateTableModelFromCsv(file, scoresTableModel)
                    }
                    refreshCardTable(scoresCard, scoresTableModel, scoresPlaceholder)

                    // Process the file (generates student_grades.xlsx)
                    val results = calculator.processFile(file)
                    if (results is List<*>) {
                        lastResults = results.filterIsInstance<StudentResult>()

                        // Load the Grades sheet from student_grades.xlsx into the grades table
                        val saved = File("student_grades.xlsx")
                        if (saved.exists()) {
                            populateTableModelFromExcel(saved, gradesTableModel, sheetName = "Grades")
                            refreshCardTable(gradesCard, gradesTableModel, gradesPlaceholder)
                        }

                        statusLbl.text      = "  ✅  Loaded ${lastResults.size} students from ${file.name}"
                        statusLbl.foreground = Theme.GREEN
                        updateResultsTable()
                        statusBar.setText("Processed ${file.name} — ${lastResults.size} students", StatusBar.StatusType.OK)
                    }
                } catch (ex: Exception) {
                    statusLbl.text      = "  ⚠  Error: ${ex.message}"
                    statusLbl.foreground = Theme.RED
                    statusBar.setText("Error: ${ex.message}", StatusBar.StatusType.ERROR)
                }
            }
        }
        return wrapper
    }

    /** Clears and repopulates a CardPanel's centre with a DarkTable, or restores the placeholder. */
    private fun refreshCardTable(card: CardPanel, model: DefaultTableModel, placeholder: JLabel) {
        card.remove(placeholder)
        // Remove any previously added scroll pane
        card.components.filterIsInstance<JScrollPane>().forEach { card.remove(it) }
        if (model.rowCount > 0) {
            card.add(darkScroll(DarkTable(model)), BorderLayout.CENTER)
        } else {
            card.add(placeholder, BorderLayout.CENTER)
        }
        card.revalidate()
        card.repaint()
    }

    /** Returns a centred muted placeholder label. */
    private fun placeholderLabel(text: String) = JLabel(text, JLabel.CENTER).also {
        it.font       = Theme.FONT_SMALL
        it.foreground = Theme.TEXT_MUTED
    }

    /** Fills a DefaultTableModel from a specific sheet (by index or name) of an Excel file. */
    private fun populateTableModelFromExcel(
        file: File,
        model: DefaultTableModel,
        sheetIndex: Int  = 0,
        sheetName: String? = null
    ) {
        model.setRowCount(0)
        model.setColumnCount(0)
        FileInputStream(file).use { fis ->
            XSSFWorkbook(fis).use { wb ->
                val sheet   = if (sheetName != null) (wb.getSheet(sheetName) ?: wb.getSheetAt(sheetIndex))
                              else wb.getSheetAt(sheetIndex)
                val maxCols = (0..sheet.lastRowNum)
                    .mapNotNull { r -> sheet.getRow(r)?.lastCellNum?.toInt() }
                    .maxOrNull() ?: 0
                for (r in 0..sheet.lastRowNum) {
                    val row = sheet.getRow(r) ?: continue
                    val cells = (0 until maxCols).map { c ->
                        when (row.getCell(c)?.cellType) {
                            CellType.NUMERIC -> "%.2f".format(row.getCell(c).numericCellValue)
                            CellType.STRING  -> row.getCell(c).stringCellValue
                            else             -> ""
                        }
                    }
                    if (r == 0) cells.forEach { model.addColumn(it) }
                    else        model.addRow(cells.toTypedArray())
                }
            }
        }
    }

    /** Fills a DefaultTableModel from a CSV file. */
    private fun populateTableModelFromCsv(file: File, model: DefaultTableModel) {
        model.setRowCount(0)
        model.setColumnCount(0)
        val rows = mutableListOf<List<String>>()
        file.bufferedReader().useLines { lines ->
            lines.filter { it.isNotBlank() }.forEach { rows.add(it.split(",").map { c -> c.trim() }) }
        }
        if (rows.isEmpty()) return
        rows.first().forEach { model.addColumn(it) }
        rows.drop(1).forEach { model.addRow(it.toTypedArray()) }
    }

    private fun buildUrlPage(): JPanel {
        val page    = darkPage()
        val wrapper = scrollWrapper(page)
        pageHeader(page, "Download URL", "Fetch and process a remote file")

        val card  = CardPanel("Remote File", "⬇")
        val inner = JPanel(GridBagLayout()).also { it.isOpaque = false }
        val gbc   = GridBagConstraints().apply { insets = Insets(6, 6, 6, 6); fill = GridBagConstraints.HORIZONTAL }

        val urlField  = StyledTextField(40)
        val dlBtn     = StyledButton("  ⬇  Download & Process  ", Theme.GREEN, Theme.BG_DEEP)
        val statusLbl = JLabel("Enter a direct link to an .xlsx or .csv file").also {
            it.foreground = Theme.TEXT_MUTED
            it.font = Theme.FONT_SMALL
        }

        gbc.gridx = 0; gbc.gridy = 0; gbc.weightx = 0.0; inner.add(lbl("File URL"), gbc)
        gbc.gridx = 1;                gbc.weightx = 1.0; inner.add(urlField, gbc)
        gbc.gridx = 0; gbc.gridy = 1; gbc.gridwidth = 2
        inner.add(JPanel(FlowLayout(FlowLayout.LEFT)).also { it.isOpaque = false; it.add(dlBtn) }, gbc)
        gbc.gridy = 2; inner.add(statusLbl, gbc)
        card.add(inner)
        page.add(padded(card, 0, 0, 0, 0))
        page.add(Box.createVerticalGlue())

        dlBtn.addActionListener {
            val url = urlField.text.trim()
            if (url.isBlank()) {
                statusLbl.foreground = Theme.RED
                statusLbl.text = "⚠  Please enter a URL."
                return@addActionListener
            }
            statusLbl.foreground = Theme.AMBER
            statusLbl.text = "Downloading…"
            statusBar.setText("Downloading…")
            SwingUtilities.invokeLater {
                try {
                    val file    = calculator.download(url)
                    val results = calculator.processFile(file)
                    if (results is List<*>) {
                        lastResults = results.filterIsInstance<StudentResult>()
                        updateResultsTable()
                        file.deleteOnExit()
                        statusLbl.foreground = Theme.GREEN
                        statusLbl.text = "✅  Processed ${lastResults.size} students."
                        statusBar.setText("Done — ${lastResults.size} students loaded", StatusBar.StatusType.OK)
                    }
                } catch (ex: Exception) {
                    statusLbl.foreground = Theme.RED
                    statusLbl.text = "⚠  ${ex.message}"
                    statusBar.setText("Download failed: ${ex.message}", StatusBar.StatusType.ERROR)
                }
            }
        }
        return wrapper
    }

    private fun buildResultsPage(): JPanel {
        val page    = darkPage()
        val wrapper = scrollWrapper(page)
        pageHeader(page, "Results", "View and export processed grade data")

        // Export toolbar
        val toolbar    = JPanel(FlowLayout(FlowLayout.LEFT, 8, 0)).also { it.isOpaque = false }
        val exportXlsx = StyledButton("  📥  Export Excel  ", Theme.GREEN, Theme.BG_DEEP)
        val exportPdf  = GhostButton("  📄  Export PDF  ",   Theme.AMBER)
        val exportHtml = GhostButton("  🌐  Export HTML  ",  Theme.ACCENT)
        toolbar.add(exportXlsx); toolbar.add(exportPdf); toolbar.add(exportHtml)
        page.add(padded(toolbar, 0, 0, 12, 0))

        // Tabbed view — Scores tab | Grades tab
        val tabbedPane = JTabbedPane().also {
            it.background = Theme.BG_PANEL
            it.foreground = Theme.TEXT_PRIMARY
            it.font       = Theme.FONT_BODY
        }
        val scoresTab = JPanel(BorderLayout()).also { it.isOpaque = false }
        val gradesTab = JPanel(BorderLayout()).also { it.isOpaque = false }
        scoresTab.add(placeholderLabel("Upload a file from 'Open File' to see scores here"), BorderLayout.CENTER)
        gradesTab.add(darkScroll(DarkTable(resultsTableModel)), BorderLayout.CENTER)
        tabbedPane.addTab("📄  Uploaded Scores", scoresTab)
        tabbedPane.addTab("✅  Generated Grades", gradesTab)
        resultsScoresTab = scoresTab

        val tabCard = CardPanel("Grade Results", "📊").also { it.layout = BorderLayout() }
        tabCard.add(tabbedPane, BorderLayout.CENTER)
        page.add(padded(tabCard, 0, 0, 0, 0))
        page.add(Box.createVerticalGlue())

        exportXlsx.addActionListener { doExport(OutputFormat.EXCEL) }
        exportPdf.addActionListener  { doExport(OutputFormat.PDF)   }
        exportHtml.addActionListener { doExport(OutputFormat.HTML)  }
        return wrapper
    }

    // ── Private helpers ───────────────────────────────────────

    private fun doExport(format: OutputFormat) {
        if (lastResults.isEmpty()) {
            statusBar.setText("No data to export", StatusBar.StatusType.ERROR)
            return
        }
        try {
            val f = calculator.generateOutput(lastResults, format)
            statusBar.setText("Exported → ${f.absolutePath}", StatusBar.StatusType.OK)
        } catch (ex: Exception) {
            statusBar.setText("Export error: ${ex.message}", StatusBar.StatusType.ERROR)
        }
    }

    private fun updateResultsTable() {
        // Refresh the grades table model
        resultsTableModel.setRowCount(0)
        resultsTableModel.setColumnCount(0)
        if (lastResults.isEmpty()) return
        val courses = lastResults.first().courseTotals.keys.toList()
        resultsTableModel.addColumn("Student")
        courses.forEach { resultsTableModel.addColumn(it); resultsTableModel.addColumn("Grade") }
        lastResults.forEach { res ->
            val row = mutableListOf<Any>(res.student)
            courses.forEach { c ->
                val total = res.courseTotals[c] ?: 0.0
                row.add("%.1f".format(total))
                row.add(calculator.calculateGrade(total))
            }
            resultsTableModel.addRow(row.toTypedArray())
        }

        // Also populate the Scores tab on the Results page from the last uploaded file
        val scoresTab = resultsScoresTab
        val uploadedFile = lastUploadedFile
        if (scoresTab != null && uploadedFile != null) {
            val scoresModel = DefaultTableModel()
            try {
                when {
                    uploadedFile.extension.equals("xlsx", true) ->
                        populateTableModelFromExcel(uploadedFile, scoresModel, sheetIndex = 0)
                    uploadedFile.extension.equals("csv", true) ->
                        populateTableModelFromCsv(uploadedFile, scoresModel)
                }
            } catch (_: Exception) {}
            scoresTab.removeAll()
            if (scoresModel.rowCount > 0) {
                scoresTab.add(darkScroll(DarkTable(scoresModel)), BorderLayout.CENTER)
            } else {
                scoresTab.add(placeholderLabel("Upload a file from 'Open File' to see scores here"), BorderLayout.CENTER)
            }
            scoresTab.revalidate()
            scoresTab.repaint()
        }

        selectPage(4)
    }

    private fun darkPage() = JPanel().also {
        it.isOpaque = false
        it.layout = BoxLayout(it, BoxLayout.Y_AXIS)
        it.background = Theme.BG_DEEP
        it.border = EmptyBorder(28, 28, 28, 28)
    }

    private fun scrollWrapper(inner: JPanel) = JPanel(BorderLayout()).also { w ->
        w.isOpaque = false
        w.add(JScrollPane(inner).also {
            it.isOpaque = false
            it.viewport.isOpaque = false
            it.border = null
        })
    }

    private fun pageHeader(page: JPanel, title: String, subtitle: String) {
        page.add(JLabel(title).also { it.font = Theme.FONT_TITLE; it.foreground = Theme.TEXT_PRIMARY; it.alignmentX = 0f })
        page.add(JLabel(subtitle).also { it.font = Theme.FONT_SMALL; it.foreground = Theme.TEXT_MUTED; it.border = EmptyBorder(2,0,0,0); it.alignmentX = 0f })
        page.add(Box.createVerticalStrut(16))
        page.add(JSeparator().also { it.foreground = Theme.BORDER; it.maximumSize = Dimension(Int.MAX_VALUE, 1) })
        page.add(Box.createVerticalStrut(20))
    }

    private fun lbl(text: String) = JLabel(text).also {
        it.font = Theme.FONT_BODY
        it.foreground = Theme.TEXT_SECONDARY
    }

    private fun padded(c: JComponent, t: Int, r: Int, b: Int, l: Int) = JPanel(BorderLayout()).also {
        it.isOpaque = false
        it.add(c)
        it.border = EmptyBorder(t, l, b, r)
        it.maximumSize = Dimension(Int.MAX_VALUE, c.preferredSize.height + t + b + 200)
    }

    private fun statCard(title: String, value: String, accent: Color, icon: String): JPanel {
        val card = object : JPanel(BorderLayout()) {
            override fun paintComponent(g: Graphics) {
                val g2 = g as Graphics2D
                g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON)
                g2.color = Theme.BG_CARD
                g2.fill(RoundRectangle2D.Float(0f, 0f, width.toFloat(), height.toFloat(), 12f, 12f))
                g2.color = accent
                g2.stroke = BasicStroke(2f)
                g2.drawLine(0, height - 2, width, height - 2)
                g2.color = Theme.BORDER
                g2.stroke = BasicStroke(1f)
                g2.draw(RoundRectangle2D.Float(0f, 0f, width - 1f, height - 1f, 12f, 12f))
                super.paintComponent(g)
            }
        }
        card.isOpaque = false
        card.border = EmptyBorder(14, 16, 14, 16)
        card.preferredSize = Dimension(0, 90)

        val iconLabel  = JLabel(icon).also { it.font = Font("SansSerif", Font.PLAIN, 22); it.foreground = accent }
        val titleLabel = JLabel(title).also { it.font = Theme.FONT_SMALL; it.foreground = Theme.TEXT_MUTED }
        val valueLabel = JLabel(value).also { it.font = Font("SansSerif", Font.BOLD, 26); it.foreground = Theme.TEXT_PRIMARY }
        val right = JPanel().also {
            it.isOpaque = false
            it.layout = BoxLayout(it, BoxLayout.Y_AXIS)
            it.add(titleLabel)
            it.add(valueLabel)
        }
        card.add(iconLabel, BorderLayout.WEST)
        card.add(right, BorderLayout.CENTER)
        return card
    }
}

// ══════════════════════════════════════════════════════════════
//  BUSINESS LOGIC
// ══════════════════════════════════════════════════════════════

enum class OutputFormat { EXCEL, PDF, HTML }

interface FileProcessor    { fun processFile(file: File): Any? }
interface OutputGenerator  { fun generateOutput(data: Any, format: OutputFormat): File }
interface Downloadable     { fun download(url: String): File }

abstract class Calculator(protected val passMark: Double = 60.0) {

    fun calculateGrade(average: Double) = when {
        average >= 90      -> "A"
        average >= 80      -> "B"
        average >= 70      -> "C"
        average >= passMark -> "D"
        else               -> "F"
    }

    fun calculateAverage(scores: List<Double>) =
        if (scores.isEmpty()) 0.0 else scores.sum() / scores.size

    fun computeAverage(scores: List<Double>, aggregator: (List<Double>) -> Double) =
        aggregator(scores)

    fun processScores(scores: List<Double>, aggregator: (List<Double>) -> Double): Pair<Double, String> {
        val avg = computeAverage(scores, aggregator)
        return avg to calculateGrade(avg)
    }

    fun processScoresWithOperation(
        scores: List<Double>,
        operation: (List<Double>) -> List<Double>,
        aggregator: (List<Double>) -> Double
    ) = aggregator(operation(scores))
}

class GradeCalculator(passMark: Double = 60.0) : Calculator(passMark), FileProcessor, OutputGenerator, Downloadable {

    override fun download(url: String): File {
        val fileName = File(URI(url).path).name.ifBlank { "downloaded.tmp" }
        val tempFile = File(fileName)
        URI(url).toURL().openStream().use { input ->
            FileOutputStream(tempFile).use { output -> input.copyTo(output) }
        }
        return tempFile
    }

    override fun processFile(file: File): Any? = when {
        file.extension.equals("xlsx", true) -> processExcel(file)
        file.extension.equals("csv",  true) -> processCsv(file)
        file.extension.equals("pdf",  true) -> emptyList<StudentResult>()
        else -> { println("Unsupported file type: ${file.extension}"); null }
    }

    override fun generateOutput(data: Any, format: OutputFormat) = when (format) {
        OutputFormat.EXCEL -> generateExcelOutput(data)
        OutputFormat.PDF   -> generatePdfOutput()
        OutputFormat.HTML  -> generateHtmlOutput(data)
    }

    private fun processExcel(file: File): List<StudentResult> {
        val results = mutableListOf<StudentResult>()
        FileInputStream(file).use { fis ->
            XSSFWorkbook(fis).use { workbook ->
                val scoresSheet = workbook.getSheet("Scores") ?: workbook.getSheetAt(0)
                workbook.getSheet("Grades")?.let { workbook.removeSheetAt(workbook.getSheetIndex(it)) }
                val gradesSheet = workbook.createSheet("Grades")
                val layout      = buildExcelLayout(scoresSheet)

                val header = gradesSheet.createRow(0)
                var colIdx = 0
                header.createCell(colIdx++).setCellValue("Student")
                for (course in layout.courseColumns.keys) {
                    header.createCell(colIdx++).setCellValue("$course Total")
                    header.createCell(colIdx++).setCellValue("$course Grade")
                }

                var outRowIdx = 1
                for (i in layout.dataStartRow..scoresSheet.lastRowNum) {
                    val row     = scoresSheet.getRow(i) ?: continue
                    val student = row.getCell(layout.nameIndex)?.toString()?.trim().orEmpty()
                    if (student.isBlank()) continue

                    val outRow      = gradesSheet.createRow(outRowIdx++)
                    var outCol      = 0
                    val totals      = mutableMapOf<String, Double>()
                    var hasAny      = false
                    outRow.createCell(outCol++).setCellValue(student)

                    for ((course, cols) in layout.courseColumns) {
                        val primary = getNumericCellValue(row.getCell(cols.first))
                        val total   = if (cols.second == null) primary else {
                            val sec = getNumericCellValue(row.getCell(cols.second!!))
                            if (primary != null && sec != null) primary + sec else null
                        }
                        if (total != null) {
                            hasAny = true
                            totals[course] = total
                            outRow.createCell(outCol++).setCellValue(total)
                            outRow.createCell(outCol++).setCellValue(calculateGrade(total))
                        } else {
                            outRow.createCell(outCol++).setCellValue("")
                            outRow.createCell(outCol++).setCellValue("")
                        }
                    }

                    if (hasAny) results.add(StudentResult(student, totals))
                    else        { gradesSheet.removeRow(outRow); outRowIdx-- }
                }

                FileOutputStream(File("student_grades.xlsx")).use { workbook.write(it) }
            }
        }
        return results
    }

    private fun processCsv(file: File): List<StudentResult> {
        val rows = mutableListOf<List<String>>()
        BufferedReader(FileReader(file)).use { reader ->
            var line = reader.readLine()
            while (line != null) {
                if (line.isNotBlank()) rows.add(parseCsvLine(line))
                line = reader.readLine()
            }
        }
        if (rows.isEmpty()) return emptyList()

        val workbook    = XSSFWorkbook()
        val layout      = buildCsvLayout(rows)
        val gradesSheet = workbook.createSheet("Grades")
        val header      = gradesSheet.createRow(0)
        var colIdx      = 0

        header.createCell(colIdx++).setCellValue("Student")
        for (course in layout.courseColumns.keys) {
            header.createCell(colIdx++).setCellValue("$course Total")
            header.createCell(colIdx++).setCellValue("$course Grade")
        }

        val results   = mutableListOf<StudentResult>()
        var outRowIdx = 1

        for (row in rows.drop(layout.dataStartRow)) {
            val student = row.getOrNull(layout.nameIndex)?.trim().orEmpty()
            if (student.isBlank()) continue

            val outRow = gradesSheet.createRow(outRowIdx++)
            var outCol = 0
            val totals = mutableMapOf<String, Double>()
            var hasAny = false
            outRow.createCell(outCol++).setCellValue(student)

            for ((course, cols) in layout.courseColumns) {
                val primary = row.getOrNull(cols.first)?.trim()?.toDoubleOrNull()
                val total   = if (cols.second == null) primary else {
                    val sec = row.getOrNull(cols.second!!)?.trim()?.toDoubleOrNull()
                    if (primary != null && sec != null) primary + sec else null
                }
                if (total != null) {
                    hasAny = true
                    totals[course] = total
                    outRow.createCell(outCol++).setCellValue(total)
                    outRow.createCell(outCol++).setCellValue(calculateGrade(total))
                } else {
                    outRow.createCell(outCol++).setCellValue("")
                    outRow.createCell(outCol++).setCellValue("")
                }
            }

            if (hasAny) results.add(StudentResult(student, totals))
            else        { gradesSheet.removeRow(outRow); outRowIdx-- }
        }

        FileOutputStream(File("student_grades.xlsx")).use { workbook.write(it) }
        workbook.close()
        return results
    }

    private fun generateExcelOutput(data: Any): File {
        val results  = data as List<StudentResult>
        val workbook = XSSFWorkbook()
        val sheet    = workbook.createSheet("Grades")
        val header   = sheet.createRow(0)
        var col      = 0

        header.createCell(col++).setCellValue("Student")
        if (results.isNotEmpty()) {
            results.first().courseTotals.keys.sorted().forEach {
                header.createCell(col++).setCellValue("$it Total")
                header.createCell(col++).setCellValue("$it Grade")
            }
        }

        results.forEachIndexed { idx, res ->
            val row = sheet.createRow(idx + 1)
            var c   = 0
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

    private fun generatePdfOutput(): File {
        val outFile = File("student_grades.pdf")
        outFile.writeText("PDF output placeholder.")
        return outFile
    }

    private fun generateHtmlOutput(data: Any): File {
        val outFile = File("student_grades.html")
        outFile.writeText(buildString {
            appendLine("<html><body><table border='1'>")
            appendLine("<tr><th>Student</th><th>Course</th><th>Total</th><th>Grade</th></tr>")
            (data as List<StudentResult>).forEach { res ->
                res.courseTotals.forEach { (course, total) ->
                    appendLine("<tr><td>${res.student}</td><td>$course</td><td>$total</td><td>${calculateGrade(total)}</td></tr>")
                }
            }
            appendLine("</table></body></html>")
        })
        return outFile
    }

    private data class ParsedLayout(
        val nameIndex: Int,
        val dataStartRow: Int,
        val courseColumns: LinkedHashMap<String, Pair<Int, Int?>>
    )

    private fun buildExcelLayout(sheet: org.apache.poi.ss.usermodel.Sheet): ParsedLayout {
        val nameIndex     = detectNameIndex(sheet)
        val coursesMarker = detectCoursesMarkerFromSheet(sheet)

        if (coursesMarker != null) {
            val (coursesRowIndex, coursesStartCol) = coursesMarker
            val courseNameRow = sheet.getRow((coursesRowIndex + 1).coerceAtMost(sheet.lastRowNum))
            val subHeaderRow  = sheet.getRow((coursesRowIndex + 2).coerceAtMost(sheet.lastRowNum))
            val hasCaExam     = subHeaderRow != null && (coursesStartCol until subHeaderRow.lastCellNum).any { col ->
                val v = subHeaderRow.getCell(col)?.toString()?.trim()?.lowercase().orEmpty()
                v == "ca" || v == "exam"
            }

            val columns     = LinkedHashMap<String, Pair<Int, Int?>>()
            var currentCourse = ""
            val maxCol = maxOf(
                courseNameRow?.lastCellNum?.toInt() ?: 0,
                subHeaderRow?.lastCellNum?.toInt()  ?: 0
            )

            for (col in coursesStartCol until maxCol) {
                val candidate = courseNameRow?.getCell(col)?.toString()?.trim().orEmpty()
                if (candidate.isNotBlank() && candidate.lowercase() !in setOf("courses", "ca", "exam")) {
                    currentCourse = candidate
                    if (!hasCaExam && col != nameIndex) columns[currentCourse] = col to null
                }
                if (!hasCaExam || currentCourse.isBlank()) continue
                when (subHeaderRow?.getCell(col)?.toString()?.trim()?.lowercase()) {
                    "ca"   -> columns[currentCourse] = col to (columns[currentCourse]?.second)
                    "exam" -> columns[currentCourse] = (columns[currentCourse]?.first ?: -1) to col
                }
            }

            if (hasCaExam) columns.entries.removeIf { (_, c) -> c.first < 0 || c.second == null }
            if (columns.isNotEmpty()) {
                val dataStart = if (hasCaExam) coursesRowIndex + 3 else coursesRowIndex + 2
                return ParsedLayout(nameIndex, dataStart, columns)
            }
        }

        val headerRowIndex = (0..sheet.lastRowNum).firstOrNull { r ->
            val row = sheet.getRow(r) ?: return@firstOrNull false
            (0 until row.lastCellNum).count { c -> !row.getCell(c)?.toString()?.trim().isNullOrBlank() } >= 2
        } ?: 0
        val header  = sheet.getRow(headerRowIndex) ?: return ParsedLayout(nameIndex, 1, linkedMapOf())
        val columns = LinkedHashMap<String, Pair<Int, Int?>>()
        for (col in 0 until header.lastCellNum) {
            if (col == nameIndex) continue
            val h = header.getCell(col)?.toString()?.trim().orEmpty()
            if (h.isNotBlank()) columns[h] = col to null
        }
        return ParsedLayout(nameIndex, headerRowIndex + 1, columns)
    }

    private fun detectNameIndex(sheet: org.apache.poi.ss.usermodel.Sheet): Int {
        for (r in 0..5.coerceAtMost(sheet.lastRowNum)) {
            val row = sheet.getRow(r) ?: continue
            for (c in 0 until row.lastCellNum) {
                val v = row.getCell(c)?.toString()?.trim()?.lowercase().orEmpty()
                if (v == "student" || v == "name") return c
            }
        }
        return 1
    }

    private fun getNumericCellValue(cell: org.apache.poi.ss.usermodel.Cell?): Double? {
        if (cell == null) return null
        return when (cell.cellType) {
            CellType.NUMERIC -> cell.numericCellValue
            CellType.STRING  -> cell.stringCellValue.toDoubleOrNull()
            else             -> null
        }
    }

    private fun parseCsvLine(line: String): List<String> {
        val result   = mutableListOf<String>()
        val sb       = StringBuilder()
        var inQuotes = false
        var i        = 0
        while (i < line.length) {
            val ch = line[i]
            when {
                ch == '"' -> {
                    if (inQuotes && i + 1 < line.length && line[i + 1] == '"') {
                        sb.append('"'); i++
                    } else {
                        inQuotes = !inQuotes
                    }
                }
                ch == ',' && !inQuotes -> { result.add(sb.toString()); sb.setLength(0) }
                else -> sb.append(ch)
            }
            i++
        }
        result.add(sb.toString())
        return result
    }

    private fun buildCsvLayout(rows: List<List<String>>): ParsedLayout {
        val nameIndex = detectNameIndexFromCsv(rows)
        val marker    = detectCoursesMarkerFromCsv(rows)

        if (marker != null) {
            val (rowIdx, startCol) = marker
            val courseNameRow      = rows.getOrNull(rowIdx + 1).orEmpty()
            val subHeaderRow       = rows.getOrNull(rowIdx + 2).orEmpty()
            val maxCol             = maxOf(courseNameRow.size, subHeaderRow.size)
            val hasCaExam          = (startCol until maxCol).any { col ->
                val v = subHeaderRow.getOrNull(col)?.trim()?.lowercase().orEmpty()
                v == "ca" || v == "exam"
            }

            val columns       = linkedMapOf<String, Pair<Int, Int?>>()
            var currentCourse = ""
            for (col in startCol until maxCol) {
                val candidate = courseNameRow.getOrNull(col)?.trim().orEmpty()
                if (candidate.isNotBlank() && candidate.lowercase() !in setOf("courses", "ca", "exam")) {
                    currentCourse = candidate
                    if (!hasCaExam && col != nameIndex) columns[currentCourse] = col to null
                }
                if (!hasCaExam || currentCourse.isBlank()) continue
                when (subHeaderRow.getOrNull(col)?.trim()?.lowercase()) {
                    "ca"   -> columns[currentCourse] = col to columns[currentCourse]?.second
                    "exam" -> columns[currentCourse] = (columns[currentCourse]?.first ?: -1) to col
                }
            }

            if (hasCaExam) columns.entries.removeIf { (_, c) -> c.first < 0 || c.second == null }
            if (columns.isNotEmpty()) {
                val dataStart = if (hasCaExam) rowIdx + 3 else rowIdx + 2
                return ParsedLayout(nameIndex, dataStart, LinkedHashMap(columns))
            }
        }

        val header  = rows.firstOrNull().orEmpty()
        val columns = linkedMapOf<String, Pair<Int, Int?>>()
        for ((col, h) in header.withIndex()) {
            if (col == nameIndex) continue
            val t = h.trim()
            if (t.isNotBlank()) columns[t] = col to null
        }
        return ParsedLayout(nameIndex, 1, LinkedHashMap(columns))
    }

    private fun detectNameIndexFromCsv(rows: List<List<String>>): Int {
        for (row in rows.take(3)) {
            for ((col, v) in row.withIndex()) {
                if (v.trim().lowercase() in setOf("student", "name")) return col
            }
        }
        return if (rows.any { it.size > 1 }) 1 else 0
    }

    private fun detectCoursesMarkerFromSheet(sheet: org.apache.poi.ss.usermodel.Sheet): Pair<Int, Int>? {
        for (r in 0..5.coerceAtMost(sheet.lastRowNum)) {
            val row = sheet.getRow(r) ?: continue
            for (c in 0 until row.lastCellNum) {
                if (row.getCell(c)?.toString()?.trim()?.lowercase() == "courses") return r to c
            }
        }
        return null
    }

    private fun detectCoursesMarkerFromCsv(rows: List<List<String>>): Pair<Int, Int>? {
        for ((r, row) in rows.take(6).withIndex()) {
            for ((c, v) in row.withIndex()) {
                if (v.trim().lowercase() == "courses") return r to c
            }
        }
        return null
    }
}

data class StudentResult(val student: String, val courseTotals: Map<String, Double>)

// ══════════════════════════════════════════════════════════════
//  ENTRY POINT
// ══════════════════════════════════════════════════════════════

fun main(args: Array<String>) {
    if (!GraphicsEnvironment.isHeadless()) {
        SwingUtilities.invokeLater { GradeCalculatorGUI().isVisible = true }
    } else {
        println("Headless environment — falling back to console mode.")
        ConsoleMenu(GradeCalculator()).run()
    }
}

// ══════════════════════════════════════════════════════════════
//  CONSOLE FALLBACK
// ══════════════════════════════════════════════════════════════

class ConsoleMenu(private val calculator: GradeCalculator) {

    private val scanner = Scanner(System.`in`)

    fun run() {
        while (true) {
            println("\n=== Student Grade Calculator ===")
            println("1. Enter scores manually")
            println("2. Process a local file (.xlsx / .csv)")
            println("3. Download and process from a URL")
            println("4. Exit")
            print("Choose: ")
            when (scanner.nextLine().trim()) {
                "1"  -> manualEntry()
                "2"  -> processLocalFile()
                "3"  -> downloadFromUrl()
                "4"  -> { println("Goodbye!"); return }
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
            print("Score $i: ")
            scores.add(scanner.nextLine().trim().toDoubleOrNull() ?: return)
        }

        val (avg, grade) = calculator.processScores(scores) { it.sum() / it.size }
        println("$name → Avg: ${"%.2f".format(avg)}, Grade: $grade")

        val bonus    = 2.0
        val avgBonus = calculator.processScoresWithOperation(scores, { l -> l.map { it + bonus } }, { l -> l.sum() / l.size })
        println("With +$bonus bonus → Avg: ${"%.2f".format(avgBonus)}, Grade: ${calculator.calculateGrade(avgBonus)}")
    }

    private fun processLocalFile() {
        print("File path: ")
        val file = File(scanner.nextLine().trim())
        if (!file.exists()) { println("File not found."); return }
        val results = calculator.processFile(file)
        if (results is List<*>) {
            val pdf  = calculator.generateOutput(results, OutputFormat.PDF)
            val html = calculator.generateOutput(results, OutputFormat.HTML)
            println("Generated PDF:  ${pdf.absolutePath}")
            println("Generated HTML: ${html.absolutePath}")
        }
    }

    private fun downloadFromUrl() {
        print("URL: ")
        try {
            val file    = calculator.download(scanner.nextLine().trim())
            val results = calculator.processFile(file)
            if (results is List<*>) {
                file.deleteOnExit()
                val pdf  = calculator.generateOutput(results, OutputFormat.PDF)
                val html = calculator.generateOutput(results, OutputFormat.HTML)
                println("Generated PDF:  ${pdf.absolutePath}")
                println("Generated HTML: ${html.absolutePath}")
            }
        } catch (e: Exception) {
            println("Failed: ${e.message}")
        }
    }
}