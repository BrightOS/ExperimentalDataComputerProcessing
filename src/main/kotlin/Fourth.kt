import kotlin.math.abs
import kotlin.math.sqrt

data class Matrix(val rows: List<List<Double>>) {
    val rowCount: Int get() = rows.size
    val colCount: Int get() = if (rows.isNotEmpty()) rows[0].size else 0

    operator fun get(i: Int, j: Int): Double = rows[i][j]
}

fun main() {
    val sourceData = readXlsFile("source_data.xls", startColumnIndex = 51, columnsCount = 10)
    val N = sourceData.regionObjects.size
    val p = sourceData.featureNames.size
    val Z = Matrix(sourceData.regionObjects.map { it.values })

    println("Z - матрица данных типа объект-признак размером Nxp\nN=$N p=$p")

    val (means, dispersions) = Z.meansAndDispersions(N, p)
    printAsTable(
        title = "Средние и дисперсии по столбцам",
        headerData = listOf("", *sourceData.featureNames.toTypedArray()),
        rows = buildList {
            add(listOf("Средние", *means.map { it.toString(accuracy = 3) }.toTypedArray()))
            add(listOf("Дисперсии", *dispersions.map { it.toString(accuracy = 3) }.toTypedArray()))
        },
    )

    val X = Z.standardize(N, p, means, dispersions)
    printAsTable(
        title = "Стандартизованная матрица",
        headerData = List(p) { "x${it + 1}" },
        tableSize = TableSize.MEDIUM,
        rows = X.rows.map { it.map { num -> num.toString(4) } }
    )

    val correl = X.correlationMatrix(N, p)
    printAsTable(
        title = "Корреляционная матрица",
        headerData = List(p) { "x${it + 1}" },
        tableSize = TableSize.MEDIUM,
        rows = correl.rows.map { it.map { num -> num.toString(4) } }
    )

    // Критерий из "замечание 1"
    var d = 0.0
    for (i in 0 until p) {
        for (j in 0 until p) {
            if (i != j) d += correl[i, j] * correl[i, j]
        }
    }
    d *= N
    val dMax = 2.009 // По таблице Стьюдента
    if (d <= dMax) {
        println("Дальнейшие вычисления нецелесообразны: $d <= $dMax")
        return
    }

    // Метод Якоби
    val (eigenVectors, eigenValuesMatrix) = correl.jacobi(p)

    // Собственные числа и сортированные собственные вектора
    val eigenValues = eigenValuesMatrix.diagonal().toMutableList()
    val (sortedEigenVectors, sortedEigenValues) = eigenVectors.sortByEigenValues(eigenValues)

    // Транспонирование
    val transposedEigenVectors = sortedEigenVectors.transpose(p)
    printAsTable(
        title = "Собственные векторы",
        headerData = List(p) { "v${it + 1}" },
        tableSize = TableSize.MEDIUM,
        rows = transposedEigenVectors.rows.map { it.map { num -> num.toString(4) } }
    )
    printAsTable(
        title = "Собственные значения",
        headerData = List(p) { "λ${it + 1}" },
        tableSize = TableSize.MEDIUM,
        rows = listOf(sortedEigenValues.map { num -> num.toString(4) })
    )

    // Проекции
    val projections = X.projectOnto(transposedEigenVectors, N, p)
    printAsTable(
        title = "Проекции объектов на новое пространство",
        headerData = List(p) { "y${it + 1}" },
        tableSize = TableSize.MEDIUM,
        rows = projections.rows.map { it.map { num -> num.toString(4) } }
    )

    // Проверка дисперсий
    val varX = X.computeVariance(N, p)
    val varY = projections.computeVariance(N, p)
    printAsTable(
        title = "Сравнение суммы дисперсий исходных признаков и проекций главных компонент",
        headerData = listOf("sum_X", "sum_Y"),
        tableSize = TableSize.SMALL,
        rows = listOf(listOf(varX.sum().toString(7), varY.sum().toString(7)))
    )

    // Минимальное p' и главные компоненты
    println()
    println()
    println()
    println("[!] Высчитываем относительную долю разброса I(p')")
    val (ip, countP) = sortedEigenValues.computeIp()
    println()
    println()
    println()
    println("[!] Минимальное p' при I(p')>0.95 = $countP")
    println("[!] Относительная доля разброса I(p') = ${ip.toString(4)}")

    // Вывод первых двух главных компонент
    printAsTable(
        title = "Главные компоненты",
        headerData = listOf("", "", "", "", "", "", "", "", "", "", "", "", ""),
        tableSize = TableSize.MEDIUM,
        rows = sortedEigenVectors.rows.indices.take(2).map { listOf("y^${it + 1}", null, null, null, null, null, null, null, null, null, null, null, sortedEigenVectors.rows[it].linearCombinationString()) },
    )

    // Тестовый пример
    val testCorrel = Matrix(
        listOf(
            listOf(1.00, 0.42, 0.54, 0.66),
            listOf(0.42, 1.00, 0.32, 0.44),
            listOf(0.54, 0.32, 1.00, 0.22),
            listOf(0.66, 0.44, 0.22, 1.00),
        )
    )
    val (testEigenVectors, testEigenValuesMatrix) = testCorrel.jacobi(4)
    val testEigenValues = testEigenValuesMatrix.diagonal()
    val (sortedTestEigenVectors, sortedTestEigenValues) = testEigenVectors.sortByEigenValues(testEigenValues)
    printAsTable(
        title = "Собственные векторы (тест)",
        headerData = List(4) { "v${it + 1}" },
        tableSize = TableSize.SMALL,
        rows = sortedTestEigenVectors.rows.map { it.map { num -> num.toString(4) } },
    )
    printAsTable(
        title = "Собственные значения (тест)",
        headerData = List(4) { "λ${it + 1}" },
        tableSize = TableSize.SMALL,
        rows = listOf(sortedTestEigenValues.map { it.toString(4) }),
    )
}

// Расширения и функции для Matrix
fun Matrix.meansAndDispersions(N: Int, p: Int): Pair<List<Double>, List<Double>> {
    val means = MutableList(p) { j -> (0 until N).sumOf { rows[it][j] } / N }
    val dispersions = MutableList(p) { j ->
        (0 until N).sumOf { (rows[it][j] - means[j]).let { it * it } } / N
    }
    return means to dispersions
}

fun Matrix.standardize(N: Int, p: Int, means: List<Double>, dispersions: List<Double>): Matrix {
    val standardized = List(N) { i ->
        List(p) { j -> (this[i, j] - means[j]) / sqrt(dispersions[j]) }
    }
    return Matrix(standardized)
}

fun Matrix.correlationMatrix(N: Int, p: Int): Matrix {
    val correl = List(p) { i ->
        List(p) { j ->
            (0 until N).sumOf { k -> this[k, i] * this[k, j] } / N
        }
    }
    return Matrix(correl)
}

fun Matrix.jacobi(n: Int): Pair<Matrix, Matrix> {
    val eps = 1e-2
    // Шаг 1. T0 = E
    val t = MutableList(n) { i -> MutableList(n) { j -> if (i == j) 1.0 else 0.0 } }
    val A = rows.map { it.toMutableList() }

    // Шаг 2. a0 - первая преграда
    val a0 = sqrt(2 * (1 until n).sumOf { j -> (0 until j).sumOf { i -> A[i][j] * A[i][j] } }) / n
    var ak = a0

    while (true) {
        // Шаг 5*. Проверяем условие выхода
        if ((1 until n).all { j -> (0 until j).all { i -> abs(A[i][j]) <= eps * a0 } }) break

        // Шаг 3. Ищем наибольший по модулю внедиагональный элемент apq > ak
        val (p, q) = (1 until n).flatMap { j ->
            (0 until j).map { i -> i to j }
        }.maxByOrNull { (i, j) -> if (abs(A[i][j]) > ak) abs(A[i][j]) else -1.0 } ?: (-1 to -1)

        // Нашли -> переходим к шагу 4
        if (p != -1) {
            // Шаг 4. Делаем вычисления
            val y = (A[p][p] - A[q][q]) / 2
            val x = if (y == 0.0) -1.0 else -sgn(y) * A[p][q] / sqrt(A[p][q] * A[p][q] + y * y)
            val s = x / sqrt(2 * (1 + sqrt(1 - x * x)))
            val c = sqrt(1 - s * s)

            for (i in 0 until n) {
                if (i != p && i != q) {
                    val z1 = A[i][p]
                    val z2 = A[i][q]
                    A[i][p] = z1 * c - z2 * s
                    A[p][i] = A[i][p]
                    A[i][q] = z1 * s + z2 * c
                    A[q][i] = A[i][q]
                }
            }
            val z5 = s * s
            val z6 = c * c
            val z7 = s * c
            val v1 = A[p][p]
            val v2 = A[p][q]
            val v3 = A[q][q]
            A[p][p] = v1 * z6 + v3 * z5 - 2 * v2 * z7
            A[q][q] = v1 * z5 + v3 * z6 + 2 * v2 * z7
            A[p][q] = (v1 - v3) * z7 + v2 * (z6 - z5)
            A[q][p] = A[p][q]

            for (i in 0 until n) {
                val z3 = t[i][p]
                val z4 = t[i][q]
                t[i][p] = z3 * c - z4 * s
                t[i][q] = z3 * s + z4 * c
            }
        }
        // Шаг 5. Находим новую преграду
        ak /= n * n
    }
    return Matrix(t) to Matrix(A)
}

fun sgn(n: Double): Int = if (n >= 0) 1 else -1

fun Matrix.sortByEigenValues(eigenValues: List<Double>): Pair<Matrix, List<Double>> {
    val sortedPairs = eigenValues
        .mapIndexed { index, value -> index to value }
        .sortedByDescending { it.second }

    val sortedIndices = sortedPairs.map { it.first }
    val sortedEigenValues = sortedPairs.map { it.second }

    val sortedRows = rows.map { row -> sortedIndices.map { row[it] } }

    return Pair(Matrix(sortedRows), sortedEigenValues)
}


fun Matrix.transpose(p: Int): Matrix {
    val transposed = List(p) { i -> List(p) { j -> this[j, i] } }
    return Matrix(transposed)
}

fun Matrix.projectOnto(t: Matrix, N: Int, p: Int): Matrix {
    val y = List(N) { j ->
        List(p) { i -> (0 until p).sumOf { k -> this[j, k] * t[i, k] } }
    }
    return Matrix(y)
}

fun Matrix.computeVariance(N: Int, p: Int): List<Double> {
    val avg = List(p) { j -> (0 until N).sumOf { i -> this[i, j] } / N }
    return List(p) { j ->
        (0 until N).sumOf { i -> (this[i, j] - avg[j]).let { it * it } } / N
    }
}

fun List<Double>.computeIp(): Pair<Double, Int> {
    val totalSum = sum()
    var cumulativeSum = 0.0
    for (i in indices) {
        println("[i] p'=$i -> I(p')=${(cumulativeSum / 10).toString(4)}")
        cumulativeSum += this[i]
        if (cumulativeSum / totalSum > 0.95) return cumulativeSum / 10 to i + 1
    }
    return 0.0 to 0
}

fun Matrix.diagonal(): List<Double> = List(rowCount) { i -> this[i, i] }

fun List<Double>.linearCombinationString(): String {
    return buildString {
        this@linearCombinationString.take(n = 10).forEachIndexed { i, value ->
            append("${value.toString(4)}*x^${i + 1}")
            if (i < 9) append(" + ")
        }
    }
}