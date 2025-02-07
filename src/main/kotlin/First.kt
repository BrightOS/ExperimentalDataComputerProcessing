import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.CellType
import org.knowm.xchart.CategoryChartBuilder
import org.knowm.xchart.SwingWrapper
import org.knowm.xchart.XYChartBuilder
import java.io.File
import java.io.FileInputStream
import kotlin.math.*

// Функция для расчета интервального ряда
private fun calculateIntervals(data: List<Double>, n: Int): List<Pair<Double, Double>> {
    // По формуле Стерджеса высчитываем количество интервалов
    val k = ceil(1 + 3.322 * log10(n.toDouble())).toInt()
    val minVal = data.minOrNull() ?: 0.0
    val maxVal = data.maxOrNull() ?: 0.0
    val intervalWidth = (maxVal - minVal) / k

    return (0 until k).map { i ->
        val lowerBound = minVal + i * intervalWidth
        val upperBound = minVal + (i + 1) * intervalWidth
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
private fun calculateMean(data: List<Double>): Double {
    return data.average()
}

// Функция для расчета выборочной дисперсии
private fun calculateDispersion(data: List<Double>, mean: Double): Double {
    return data.fold(0.0) { acc, x -> acc + (x - mean).pow(2) } / (data.size - 1)
}

// Функция для расчета выборочного среднеквадратического отклонения
private fun calculateStdDev(variance: Double): Double {
    return sqrt(variance)
}

// Функция для расчета интервальной оценки математического ожидания
private fun calculateConfidenceInterval(
    mean: Double,
    stdDev: Double,
    n: Int,
): Pair<Double, Double> {
    val tValue = 1.96 // Для alpha = 0.05 и больших n
    val marginOfError = tValue * (stdDev / sqrt(n.toDouble()))
    return Pair(mean - marginOfError, mean + marginOfError)
}

// Функция для расчета относительной точности оценки математического ожидания
private fun calculateRelativeAccuracy(confidenceInterval: Pair<Double, Double>, mean: Double): Double {
    val width = confidenceInterval.second - confidenceInterval.first
    return width / (2 * mean)
}

// Функция для расчета размаха вариации
private fun calculateRange(data: List<Double>): Double {
    return (data.maxOrNull() ?: 0.0) - (data.minOrNull() ?: 0.0)
}

// Функция для расчета коэффициента вариации
private fun calculateCoefficientOfVariation(stdDev: Double, mean: Double): Double {
    return stdDev / mean
}

fun readXlsFile(filePath: String): List<TableEntry> {
    val data = mutableListOf<TableEntry>()
    val columnIndex = 53

    FileInputStream(File(filePath)).use { fis ->
        val workbook = HSSFWorkbook(fis)
        val sheet: HSSFSheet = workbook.getSheetAt(0)
        var rowIndex = 2

        println("Выбран столбец с данными следующего типа:\n${sheet.getRow(0).getCell(columnIndex)}\n")

        while (data.size < 50) {
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
                println("Отсутствует значение у региона. Пропускаем $rowName")
            }
        }
    }

    println(data.map { it.cellValue })

    return data
}


// Основная функция
fun main() {
    // Пример данных (замените их на свои данные из лабораторной работы)
    val data = readXlsFile("source_data.xls").map { it.cellValue }
    val n = data.size

    // 1. Построение интервального ряда
    val intervals = calculateIntervals(data, n)
    val frequencies = calculateFrequencies(data, intervals)
    val relativeFrequencies = frequencies.map { it.toDouble() / n }
    val cumulativeFrequencies = frequencies.runningFold(0.0) { acc, freq -> acc + freq }.drop(1)

    // 2. Расчет статистических характеристик
    val mean = calculateMean(data)
    val dispersion = calculateDispersion(data, mean)
    val stdDev = calculateStdDev(dispersion)
    val confidenceInterval = calculateConfidenceInterval(mean, stdDev, n)
    val relativeAccuracy = calculateRelativeAccuracy(confidenceInterval, mean)
    val range = calculateRange(data)
    val coefficientOfVariation = calculateCoefficientOfVariation(stdDev, mean)

    // 3. Вывод результатов
    println("Интервалы:")
    intervals.forEachIndexed { index, interval ->
        println("Интервал ${index + 1}: [${interval.first}, ${interval.second}]")
    }

    println("\nЧастоты: $frequencies")
    println("Относительные частоты: $relativeFrequencies")
    println("Кумулятивные частоты: $cumulativeFrequencies")

    println("\nВыборочное среднее: $mean")
    println("Выборочная дисперсия: $dispersion")
    println("Выборочное среднеквадратическое отклонение: $stdDev")
    println("Интервальная оценка математического ожидания: $confidenceInterval")
    println("Относительная точность оценки математического ожидания: $relativeAccuracy")
    println("Размах вариации: $range")
    println("Коэффициент вариации: $coefficientOfVariation")

    // 4. Анализ результатов
    println("\nАнализ результатов:")
    println("На основе анализа можно сделать выводы о распределении данных и их вариативности.")

    // Построение гистограммы
    val histogramChart = CategoryChartBuilder()
        .width(800)
        .height(600)
        .title("Гистограмма")
        .xAxisTitle("Интервалы")
        .yAxisTitle("Частоты")
        .build()

    histogramChart.styler.isLegendVisible = false
    histogramChart.addSeries("Частоты", intervals.map { it.first }, frequencies)

    // Середины интервалов для полигона
    val midpoints = intervals.map { (it.first + it.second) / 2 }

    // Построение полигона
    val polygonChart = XYChartBuilder()
        .width(800)
        .height(600)
        .title("Полигон")
        .xAxisTitle("Середины интервалов")
        .yAxisTitle("Относительные частоты")
        .build()

    polygonChart.styler.isLegendVisible = false
    polygonChart.addSeries("Полигон", midpoints, relativeFrequencies)

    // Построение кумуляты
    val cumulativeChart = CategoryChartBuilder()
        .width(800)
        .height(600)
        .title("Кумулята")
        .xAxisTitle("Интервалы")
        .yAxisTitle("Накопленные частоты")
        .build()

    cumulativeChart.styler.isLegendVisible = false
    cumulativeChart.addSeries("Кумулята", intervals.map { it.first }, cumulativeFrequencies)

    // Отображение графиков
    SwingWrapper(histogramChart).displayChart()
    SwingWrapper(polygonChart).displayChart()
    SwingWrapper(cumulativeChart).displayChart()
}

data class TableEntry(
    val rowName: String,
    val cellValue: Double,
)