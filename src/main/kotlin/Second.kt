import kotlin.math.pow
import kotlin.math.sqrt

fun main() {
    val sourceData = readXlsFile(
        filePath = "source_data.xls",
        startColumnIndex = 51,
        columnsCount = 10,
    )
    correlationAnalysis(
        forSecondLab = true,
        sourceData = sourceData,
    )
}

fun correlationAnalysis(
    forSecondLab: Boolean,
    sourceData: SourceData,
): Pair<Int, List<Int>> {
    val data = sourceData.regionObjects.map { it.values }

    val objectsCount = data.size
    val featuresCount = data[0].size
    val featureIndicesArray = sourceData.featureNames.indices.map { it.toString() }.toTypedArray()

    // 1a) Средние и дисперсии по столбцам
    if (forSecondLab) println()
    val means = DoubleArray(featuresCount) { j -> data.sumOf { it[j] } / objectsCount }
    val dispersions = DoubleArray(featuresCount) { j -> data.sumOf { (it[j] - means[j]).pow(2) } / objectsCount }
    if (forSecondLab) printAsTable(
        title = "[1А] Средние и дисперсии по столбцам",
        headerData = listOf("", *sourceData.featureNames.toTypedArray()),
        rows = buildList {
            add(listOf("Средние", *means.map { it.toString(accuracy = 3) }.toTypedArray()))
            add(listOf("Дисперсии", *dispersions.map { it.toString(accuracy = 3) }.toTypedArray()))
        },
    )

    // 1б) Стандартизированная матрица
    if (forSecondLab) println()
    val X = data.map { row -> row.mapIndexed { j, value -> (value - means[j]) / sqrt(dispersions[j]) } }
    if (forSecondLab) printAsTable(
        title = "[1Б] Стандартизованная матрица",
        headerData = listOf("Регион", *sourceData.featureNames.toTypedArray()),
        rows = X.mapIndexed { index, it ->
            listOf(sourceData.regionObjects[index].name, *it.map { it.toString(accuracy = 3) }.toTypedArray())
        },
    )
    if (forSecondLab) println()

    // 1в) Ковариационная матрица
    val covar =
        Array(featuresCount) { i -> DoubleArray(featuresCount) { j -> data.sumOf { (it[i] - means[i]) * (it[j] - means[j]) } / objectsCount } }
    if (forSecondLab) printAsTable(
        title = "[1В] Ковариационная матрица",
        headerData = listOf("Признак", *featureIndicesArray),
        rows = covar.mapIndexed { index, it ->
            listOf(index.toString(), *it.map { it.toString(accuracy = 2) }.toTypedArray())
        },
        tableSize = TableSize.MEDIUM,
    )
    if (forSecondLab) println()

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
            if (i != j) (corel[i][j] * sqrt(objectsCount - 2.0)) / sqrt(1 - corel[i][j].pow(2)) else Double.NaN
        }
    }
    if (forSecondLab) printAsTable(
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
    if (forSecondLab) println()

    val alpha = 0.05
    val f = objectsCount - 2
    val tTable = 1.993
    if (forSecondLab) {
        println()
        println()
        println("[i] alpha (приемлемая вероятность ошибки) = $alpha")
        println("[i] f = n - 2 = $f")
        println("[i] tтабл (по таблице t-критерия Стьюдента) = $tTable")
    }

    val resultMatrix = T.map { it.map { if (it.isNaN()) "[ ]" else if (kotlin.math.abs(it) >= tTable) "H1" else "H0" } }

    printAsTable(
        title = "Решение о гипотезах",
        headerData = listOf("Признак", *featureIndicesArray),
        rows = resultMatrix.mapIndexed { index, it ->
            listOf(
                index.toString(),
                *it.toTypedArray()
            )
        },
        tableSize = TableSize.SMALL
    )
    println()

    val counted = resultMatrix.map { strings ->
        strings.map { it.contains("H1") }
    }
    val max = counted.withIndex().maxBy { it.value.count { it } }
    val xIndices = max.value.mapIndexed { index, b -> if (b) index else null }.filterNotNull()
    val yIndex = max.index

    if (!forSecondLab) {
        println()
        println()
        println("[i] Выбираем в качестве наблюдаемого признака: ${yIndex}")
        println("[i] В качестве зависимых признаков: ${xIndices}")
        println()
    }

    return yIndex to xIndices
}
