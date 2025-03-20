import java.util.Scanner

fun main() {
    val scanner = Scanner(System.`in`)

    val X = Array(size = 5) { DoubleArray(size = 3) { 1.0 } }
    val y = DoubleArray(size = 5) { .0 }

    println("[!] Необходимо ввести данные матрицы зависимых признаков X (5x2) (в каждой строке по 2 элемента):")
    repeat(5) {
        print("> ")
        X[it][0] = scanner.nextInt().toDouble()
        X[it][1] = scanner.nextInt().toDouble()
    }
    println()

    println("[!] Необходимо ввести данные матрицы вектора y (5x1) (в одной строке 5 элементов)")
    print("> ")
    repeat(5) {
        y[it] = scanner.nextInt().toDouble()
    }

    regressionAnalysis(X = X, y = y)
}