import de.vandermeer.asciitable.AT_Context
import de.vandermeer.asciitable.AsciiTable
import de.vandermeer.skb.interfaces.transformers.textformat.TextAlignment
import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import java.io.File
import java.io.FileInputStream

private const val COLUMNS_COUNT_FOR_FILTER = 132

data class SourceData(
    val featureNames: List<String>,
    val regionObjects: List<RegionObject>
)

data class RegionObject(
    val name: String,
    val values: List<Double>,
)

@Suppress("SameParameterValue")
fun readXlsFile(
    filePath: String,
    startColumnIndex: Int,
    columnsCount: Int,
): SourceData {
    val featureNames = mutableListOf<String>()
    val regionObjects = mutableListOf<RegionObject>()

    val columnIndices = (startColumnIndex until startColumnIndex + columnsCount)

    FileInputStream(File(filePath)).use { fis ->
        val workbook = HSSFWorkbook(fis)
        val sheet: HSSFSheet = workbook.getSheetAt(0)

        featureNames.addAll(columnIndices.map { sheet.getRow(0).getCell(it).stringCellValue })

        for (row in sheet.drop(2).take(78)) {
            val values = row.map { it }
                .slice(columnIndices)
                .map { cell ->
                    runCatching { cell.numericCellValue }.getOrNull()
                }

            if (values.isNotEmpty() && !values.contains(null)) {
                regionObjects.add(
                    RegionObject(
                        name = row.getCell(0).stringCellValue,
                        values = values.filterNotNull(),
                    )
                )
            }
        }
    }

    println("[i] Количество объектов (N) = ${regionObjects.size}")
    println("[i] Количество признаков (p) = $columnsCount")
    printAsTable(
        title = "[i] Матрица типа \"Объект - признак\" размером ${regionObjects.size} * $columnsCount (N * p)",
        headerData = listOf("Регион", *featureNames.toTypedArray()),
        rows = buildList {
            regionObjects.forEach {
                add(listOf(it.name, *it.values.map { it.toString(accuracy = 1) }.toTypedArray()))
            }
        },
    )

    return SourceData(
        featureNames = featureNames,
        regionObjects = regionObjects,
    )
}

enum class TableSize {
    SMALL, MEDIUM, LARGE
}

fun printAsTable(
    title: String? = null,
    headerData: List<String?>,
    rows: List<List<String?>>,
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
//                    this.setPaddingTopBottom(1)
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