import kotlin.math.pow

fun main() {
    val data = readXlsFile(
        filePath = "source_data.xls",
        startColumnIndex = 51,
        columnsCount = 10,
    )

    val (yIndex, xIndices) = correlationAnalysis(
        forSecondLab = false,
        sourceData = data,
    )

    val X = data.regionObjects
        .map { it.values.filterIndexed { index, _ -> index in xIndices }.plus(1.0).toDoubleArray() }
        .toTypedArray()

    val y = data.regionObjects.map { it.values[yIndex] }.toDoubleArray()

    regressionAnalysis(X = X, y = y)
}

fun multiplyMatrices(A: Array<DoubleArray>, B: Array<DoubleArray>): Array<DoubleArray> {
    val result = Array(A.size) { DoubleArray(B[0].size) }
    for (i in A.indices) {
        for (j in B[0].indices) {
            for (k in B.indices) {
                result[i][j] += A[i][k] * B[k][j]
            }
        }
    }
    return result
}

fun transposeMatrix(matrix: Array<DoubleArray>): Array<DoubleArray> {
    val rows = matrix.size
    val cols = matrix[0].size
    val transposed = Array(cols) { DoubleArray(rows) }
    for (i in 0 until rows) {
        for (j in 0 until cols) {
            transposed[j][i] = matrix[i][j]
        }
    }
    return transposed
}

// Инвертирование матрицы методом Жордана-Гаусса
fun inverseMatrix(matrix: Array<DoubleArray>): Array<DoubleArray> {
    val n = matrix.size
    val identity = Array(n) { i -> DoubleArray(n) { if (it == i) 1.0 else 0.0 } }
    val augmented = Array(n) { i -> matrix[i] + identity[i] }

    for (i in 0 until n) {
        val diag = augmented[i][i]
        for (j in 0 until 2 * n) augmented[i][j] /= diag
        for (k in 0 until n) {
            if (k != i) {
                val factor = augmented[k][i]
                for (j in 0 until 2 * n) augmented[k][j] -= factor * augmented[i][j]
            }
        }
    }
    return Array(n) { i -> augmented[i].sliceArray(n until 2 * n) }
}

fun regressionAnalysis(
    X: Array<DoubleArray>,
    y: DoubleArray,
) {
    printAsTable(
        headerData = listOf("y", *Array(X.first().size) { "X${it + 1}" }),
        rows = X.mapIndexed { index, it ->
            listOf(y[index].toString(accuracy = 1)).plus(it.map { it.toString(accuracy = 1) }.toList())
        },
        tableSize = TableSize.SMALL,
    )

    val XT = transposeMatrix(X)
    val XTX = multiplyMatrices(XT, X)
    val XTy = multiplyMatrices(XT, Array(y.size) { doubleArrayOf(y[it]) })
    val XTX_inv = inverseMatrix(XTX)
    val a = multiplyMatrices(XTX_inv, XTy)

    val yPred = multiplyMatrices(X, a).map { it[0] }.toDoubleArray()
    val yMean = y.average()
    val yPredMean = yPred.average()
    val r2 = 1 - y.indices.sumOf { (y[it] - yPred[it]).pow(2) } / y.sumOf { (it - yMean).pow(2) }

    printAsTable(
        title = "Коэффициенты регрессии",
        headerData = listOf("№", "Коэффициент"),
        rows = a.mapIndexed { index, value -> listOf((index + 1).toString(), value[0].toString(accuracy = 10)) },
        tableSize = TableSize.SMALL
    )

    printAsTable(
        title = "Средние значения ground-truth и предсказанных значений",
        headerData = listOf("Ground-truth Mean", "Predicted Mean"),
        rows = listOf(listOf(yMean.toString(accuracy = 15), yPredMean.toString(accuracy = 15))),
        tableSize = TableSize.SMALL
    )

    printAsTable(
        title = "Коэффициент детерминации",
        headerData = listOf("R^2"),
        rows = listOf(listOf(r2.toString(accuracy = 4))),
        tableSize = TableSize.SMALL
    )
}