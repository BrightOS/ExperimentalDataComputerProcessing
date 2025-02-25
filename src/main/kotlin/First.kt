import de.vandermeer.asciitable.AT_Context
import de.vandermeer.asciitable.AsciiTable
import de.vandermeer.skb.interfaces.transformers.textformat.TextAlignment
import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.CellType
import org.knowm.xchart.CategoryChartBuilder
import org.knowm.xchart.CategorySeries
import org.knowm.xchart.SwingWrapper
import org.knowm.xchart.XYChartBuilder
import java.io.File
import java.io.FileInputStream
import kotlin.math.*

// Функция для расчета интегральной функции распределения
private fun calculateCumulativeDistribution(relativeFrequencies: List<Double>): List<Double> {
    return relativeFrequencies.runningFold(0.0) { acc, freq -> acc + freq }.drop(1)
}

// Функция для расчета дифференциальной функции распределения
private fun calculateDensityFunction(relativeFrequencies: List<Double>, intervalWidth: Double): List<Double> {
    return relativeFrequencies.map { it / intervalWidth }
}

// Функция для расчета интервального ряда
private fun calculateIntervals(data: List<Double>, n: Int): List<Pair<Double, Double>> {
    // По формуле Стерджеса высчитываем количество интервалов
    val k = 1 + 3.322 * log10(n.toDouble())
    val minVal = floor(data.minOrNull() ?: 0.0)
    val maxVal = ceil(data.maxOrNull() ?: 0.0)
    val intervalWidth = (maxVal - minVal) / k
    println()
    println("[i] Количество интервалов:")
    println("k = 1 + 3.322 * lg($n) = $k")
    val ceilK = ceil(k).toInt()
    println("Принимаем k = $ceilK")
    println()
    println("[i] Ширина интервала:")
    println("Δl = data.max() - data.min() / k = (${data.maxOrNull() ?: 0.0} - ${data.minOrNull() ?: 0.0}) / $k = $intervalWidth")
    val ceilIntervalWidth = round(intervalWidth).toInt()
    println("Принимаем Δl = $ceilIntervalWidth")

    return (0 until ceilK).map { i ->
        val lowerBound = minVal + i * ceilIntervalWidth
        val upperBound = minVal + (i + 1) * ceilIntervalWidth
        Pair(lowerBound, upperBound)
    }
}

// Подсчитывание частот
private fun calculateFrequencies(data: List<Double>, intervals: List<Pair<Double, Double>>): List<Int> {
    return intervals.map { (lower, upper) ->
        data.count { it in lower..upper }
    }
}

// Функция для расчета выборочного среднего
private fun calculateMean(midpoints: List<Double>, relativeFrequencies: List<Double>): Double {
    var sum = .0
    println()
    println("[i] Выборочное среднее:")
    print("l = ")
    repeat(midpoints.size) {
        sum += midpoints[it] * relativeFrequencies[it]
        print("${midpoints[it]} * ${relativeFrequencies[it]} ")
        if (it < midpoints.size - 1) print("+ ")
    }
    println("= $sum")
    return sum
}

// Функция для расчета выборочной дисперсии
private fun calculateDispersion(
    size: Int,
    midpoints: List<Double>,
    frequencies: List<Int>,
    mean: Double,
): Double {
    println()
    println("[i] Дисперсия:")
    print("D = 1 / N * Σ[(li - l)^2 * ni] = 1 / $size * [(${midpoints.first()} - $mean)^2 * ${frequencies.first()} + ... + (${midpoints.last() - mean})^2 * ${frequencies.last()}]")
    var sum = .0
    repeat(midpoints.size) {
        sum += (midpoints[it] - mean).pow(2) * frequencies[it]
    }
    sum /= size
    println(" = $sum")
    return sum
}

// Функция для расчета выборочного среднеквадратического отклонения
private fun calculateStdDev(variance: Double): Double {
    println()
    println("[i] Выборочное среднеквадратическое отклонение:")
    val stdDev = sqrt(variance)
    println("σ = sqrt(D) = sqrt($variance) = $stdDev")
    return stdDev
}

// Функция для расчета интервальной оценки математического ожидания
private fun calculateConfidenceInterval(
    size: Int,
    mean: Double,
    stdDev: Double,
    n: Int,
): Pair<Double, Double> {
    println()
    println("[i] Ширина доверительного интервала:")
    val tValue = 1.67
    println("Для α = 0.05 и n = 50: t = $tValue")
    print("Δ = t * σ / sqrt(size - 1) = $tValue * $stdDev / sqrt(${size - 1}) = ")
    val marginOfError = tValue * stdDev / sqrt(n - 1.0)
    println(marginOfError)
    val result = Pair(mean - marginOfError, mean + marginOfError)
    println()
    println("[i] Доверительный интервал:")
    println("($mean - $marginOfError) < M(l) < ($mean + $marginOfError)")
    println("Т.е. ${result.first} < M(L) < ${result.second}.")
    println()
    println("С вероятностью p = 0.95 математическое ожидание будет находиться в интервале\nот ${result.first} до ${result.second}. Только у 5% значений генеральной\nсовокупности будут иметь мат. ожидание вне интервала.")
    return result
}

// Функция для расчета относительной точности оценки математического ожидания
private fun calculateRelativeAccuracy(confidenceInterval: Pair<Double, Double>, mean: Double): Double {
    println()
    println("[i] Относительная точность оценки математического ожидания:")
    val width = confidenceInterval.second - confidenceInterval.first
    val result = width / (2 * mean)
    println("μ = Δ / l = ${width / 2} / $mean = $result")
    println()
    println("Это значит, что с доверительной вероятностью pd = 0.95 относительная ошибка при\nоценке среднего значения не превысит ±${String.format("%.2f", result * 100)}% от величины среднего значения")
    return result
}

// Функция для расчета размаха вариации
private fun calculateRange(data: List<Double>): Double {
    println()
    println("[i] Размах вариации результатов эксперимента:")
    val maxValue = (data.maxOrNull() ?: 0.0)
    val minValue = (data.minOrNull() ?: 0.0)
    val result = maxValue - minValue
    println("w = data.max() - data.min() = $maxValue - $minValue = $result")
    return result
}

// Функция для расчета коэффициента вариации
private fun calculateCoefficientOfVariation(stdDev: Double, mean: Double): Double {
    println()
    println("[i] Коэффициент вариации:")
    val result = stdDev / mean
    println("ν = σ / l = $stdDev / $mean = $result")
    return result
}

fun readXlsFile(filePath: String, columnIndex: Int): List<TableEntry> {
    val data = mutableListOf<TableEntry>()

    FileInputStream(File(filePath)).use { fis ->
        val workbook = HSSFWorkbook(fis)
        val sheet: HSSFSheet = workbook.getSheetAt(0)
        var rowIndex = 2

        println("[i] ${sheet.getRow(0).getCell(columnIndex)}\n")

        while (data.size < 55) {
            val row = sheet.getRow(rowIndex++)
            val rowName = row.getCell(0).stringCellValue

            val cell = row.getCell(columnIndex)
            if (cell?.cellType == CellType.NUMERIC) {
                data.add(
                    TableEntry(
                        rowName = rowName,
                        cellValue = cell.numericCellValue,
                    )
                )
            } else {
                println("[-] $rowName (empty)")
            }
        }
    }

    val sortedData = data.sortedByDescending { it.cellValue }
    sortedData.take(5).forEach {
        println("[-] ${it.rowName} (too big)")
    }

    val resultData = sortedData.drop(5)

    println()
    println("[i] Выбранные данные:")
    println(resultData.map { it.cellValue }.joinToString(" "))

    return resultData
}


// Основная функция
fun main() {
    // Пример данных (замените их на свои данные из лабораторной работы)
    val data = readXlsFile(filePath = "source_data.xls", columnIndex = 52).map { it.cellValue }
    val n = data.size

    val intervals = calculateIntervals(data, n)
    val midpoints = intervals.map { (it.first + it.second) / 2 }
    val intervalsCount = intervals.size
    val frequencies = calculateFrequencies(data, intervals)
    val relativeFrequencies = frequencies.map { it.toDouble() / n }
    val cumulativeFrequencies = frequencies.runningFold(0.0) { acc, freq -> acc + freq }.drop(1)
    val cumulativeDistribution = calculateCumulativeDistribution(relativeFrequencies)
    val densityFunction = calculateDensityFunction(relativeFrequencies, 7.0)

    println()
    val table = AsciiTable(AT_Context().apply {
        setWidth(200)
    }).apply {
        addRule()
        addRow("Наименование параметра", "Обозначение", *Array(intervalsCount) { "k${it + 1}" }).apply {
            this.setTextAlignment(TextAlignment.CENTER)
        }
        addRule().apply {
        }
        addRow("Границы интервалов", "(ci-1,ci)", *intervals.toTypedArray())
        addRule()
        addRow("Середины интервалов", "li", *midpoints.toTypedArray())
        addRule()
        addRow("Частота", "ni", *frequencies.toTypedArray())
        addRule()
        addRow("Относительная частота", "mi", *relativeFrequencies.toTypedArray())
        addRule()
        addRow("Накопленная частота", "sum(nj)", *cumulativeFrequencies.toTypedArray())
        addRule()
        addRow("Оценка интегр ф-ии", "F(ci)", *cumulativeDistribution.toTypedArray())
        addRule()
        addRow("Оценка дифф ф-ии", "f(li)", *densityFunction.toTypedArray())
        addRule()
    }
    println(table.render())

    val mean = calculateMean(midpoints, relativeFrequencies)
    val dispersion = calculateDispersion(data.size, midpoints, frequencies, mean)
    val stdDev = calculateStdDev(dispersion)
    val confidenceInterval = calculateConfidenceInterval(data.size, mean, stdDev, n)
    calculateRelativeAccuracy(confidenceInterval, mean)
    calculateRange(data)
    calculateCoefficientOfVariation(stdDev, mean)

    // Построение гистограммы
    val histogramChart = CategoryChartBuilder()
        .width(800)
        .height(600)
        .title("Гистограмма")
        .xAxisTitle("Интервалы")
        .yAxisTitle("Частоты")
        .build()

    histogramChart.styler.isLegendVisible = false
    histogramChart.styler.setOverlapped(true)
    histogramChart.addSeries("Частоты", midpoints, densityFunction)
    val overlappedLine: CategorySeries = histogramChart.addSeries("Полигон", midpoints, densityFunction)
    overlappedLine.setChartCategorySeriesRenderStyle(CategorySeries.CategorySeriesRenderStyle.Line)

    // Построение кумуляты
    val cumulativeChart = XYChartBuilder()
        .width(800)
        .height(600)
        .title("Кумулята")
        .xAxisTitle("Интервалы")
        .yAxisTitle("Накопленные частоты")
        .build()

    val cumulativeChartXData = listOf(.0).plus(
        intervals.flatMapIndexed { index: Int, pair: Pair<Double, Double> ->
            listOf(
                if (index == 0) 0 else pair.first - (pair.second - pair.first),
                pair.first
            )
        }
    )

    val cumulativeChartYData = listOf(.0).plus(
        cumulativeDistribution.flatMapIndexed { index: Int, d: Double ->
            listOf(
                cumulativeDistribution[index],
                cumulativeDistribution[index]
            )
        }
    )

    cumulativeChart.styler.isLegendVisible = false
    cumulativeChart.addSeries(
        "Кумулята",
        cumulativeChartXData,
        cumulativeChartYData,
    )

    // Отображение графиков
    SwingWrapper(histogramChart).displayChart()
    SwingWrapper(cumulativeChart).displayChart()
}

data class TableEntry(
    val rowName: String,
    val cellValue: Double,
)