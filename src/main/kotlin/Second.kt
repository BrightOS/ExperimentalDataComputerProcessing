import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import java.io.File
import java.io.FileInputStream
import kotlin.math.pow
import kotlin.math.sqrt

data class SourceData(
    val featureNames: List<String>,
    val regionObjects: List<RegionObject>
)

data class RegionObject(
    val name: String,
    val values: List<Double>,
)

@Suppress("SameParameterValue")
private fun readXlsFile(
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
            val values = row.toList()
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
    }

    return SourceData(
        featureNames = featureNames,
        regionObjects = regionObjects,
    )
}

fun main() {
    val sourceData = readXlsFile(
        filePath = "source_data.xls",
        startColumnIndex = 51,
        columnsCount = 10,
    )

    val data = sourceData.regionObjects.map { it.values }

    val objectsCount = data.size
    val featuresCount = data[0].size
    val featureIndicesArray = sourceData.featureNames.indices.map { it.toString() }.toTypedArray()

    // 1a) Средние и дисперсии по столбцам
    println()
    val means = DoubleArray(featuresCount) { j -> data.sumOf { it[j] } / objectsCount }
    val dispersions = DoubleArray(featuresCount) { j -> data.sumOf { (it[j] - means[j]).pow(2) } / objectsCount }
    printAsTable(
        title = "[1А] Средние и дисперсии по столбцам",
        headerData = listOf("", *sourceData.featureNames.toTypedArray()),
        rows = buildList {
            add(listOf("Средние", *means.map { it.toString(accuracy = 3) }.toTypedArray()))
            add(listOf("Дисперсии", *dispersions.map { it.toString(accuracy = 3) }.toTypedArray()))
        },
    )

    // 1б) Стандартизированная матрица
    println()
    val X = data.map { row -> row.mapIndexed { j, value -> (value - means[j]) / sqrt(dispersions[j]) } }
    printAsTable(
        title = "[1Б] Стандартизованная матрица",
        headerData = listOf("Регион", *sourceData.featureNames.toTypedArray()),
        rows = X.mapIndexed { index, it ->
            listOf(sourceData.regionObjects[index].name, *it.map { it.toString(accuracy = 3) }.toTypedArray())
        },
    )
    println()

    // 1в) Ковариационная матрица
    val covar =
        Array(featuresCount) { i -> DoubleArray(featuresCount) { j -> data.sumOf { (it[i] - means[i]) * (it[j] - means[j]) } / objectsCount } }
    printAsTable(
        title = "[1В] Ковариационная матрица",
        headerData = listOf("Признак", *featureIndicesArray),
        rows = covar.mapIndexed { index, it ->
            listOf(index.toString(), *it.map { it.toString(accuracy = 2) }.toTypedArray())
        },
        tableSize = TableSize.MEDIUM,
    )
    println()

    // 1г) Корреляционная матрица
    val corel =
        Array(featuresCount) { i -> DoubleArray(featuresCount) { j -> X.sumOf { it[i] * it[j] } / objectsCount } }
    printAsTable(
        title = "[1В] Корреляционная матрица",
        headerData = listOf("Признак", *featureIndicesArray),
        rows = corel.mapIndexed { index, it ->
            listOf(index.toString(), *it.map { it.toString(accuracy = 3) }.toTypedArray())
        },
        tableSize = TableSize.MEDIUM,
    )
    println()

    // 2) Проверка гипотезы о значимости коэффициентов корреляции
    val T = Array(featuresCount) { i ->
        DoubleArray(featuresCount) { j ->
            if (i != j) (corel[i][j] * sqrt(objectsCount - 2.0)) / sqrt(1 - corel[i][j] * corel[i][j]) else Double.NaN
        }
    }
    printAsTable(
        title = "[2] Проверка гипотезы о значимости коэффициентов корреляции",
        headerData = listOf("Признак", *featureIndicesArray),
        rows = T.mapIndexed { index, it ->
            listOf(
                index.toString(),
                *it.map { if (it.isNaN()) "[ ]" else it.toString(accuracy = 3) }.toTypedArray()
            )
        },
        tableSize = TableSize.MEDIUM,
    )
    println()

    val alpha = 0.05
    val f = objectsCount - 2
    val tTable = 2.0024655
    println("alpha=$alpha, f=$f, t_table=$tTable")
    println()

    printAsTable(
        title = "Решение о гипотезах",
        headerData = listOf("Признак", *featureIndicesArray),
        rows = T.mapIndexed { index, it ->
            listOf(
                index.toString(),
                *it.map { if (it.isNaN()) "[ ]" else if (kotlin.math.abs(it) >= tTable) "H1" else "H0" }.toTypedArray()
            )
        },
        tableSize = TableSize.SMALL
    )
}
