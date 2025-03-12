import de.vandermeer.asciitable.AT_Context
import de.vandermeer.asciitable.AsciiTable
import de.vandermeer.skb.interfaces.transformers.textformat.TextAlignment

enum class TableSize {
    SMALL, MEDIUM, LARGE
}

fun printAsTable(
    title: String? = null,
    headerData: List<String>,
    rows: List<List<String>>,
    tableSize: TableSize = TableSize.LARGE,
) {
    println("\n\n")
    println(
        AsciiTable(AT_Context().apply {
            when (tableSize) {
                TableSize.SMALL -> Unit
                TableSize.MEDIUM -> setWidth(150)
                TableSize.LARGE -> setWidth(200)
            }
        }).apply {
            if (title != null) {
                addRule()
                addRow(*Array<String?>(headerData.size - 1) { null }, title).apply {
                    this.setTextAlignment(TextAlignment.CENTER)
                    this.setPaddingTopBottom(1)
                }
            }
            addRule()
            addRow(*headerData.toTypedArray()).apply {
                this.setTextAlignment(TextAlignment.CENTER)
            }
            addRule()
            rows.forEach {
                addRow(*it.toTypedArray()).apply {
                    this.setTextAlignment(TextAlignment.CENTER)
                }
                addRule()
            }
        }.render()
    )
}

fun Double.toString(accuracy: Int = 2): String = String.format("%.${accuracy}f", this)