import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.sikuli.script.Screen
import java.awt.Desktop
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.RandomAccessFile
import javax.swing.JButton
import javax.swing.JFrame
import javax.swing.JPanel
import javax.swing.SwingUtilities


fun main() {
    // Путь к Вашему файлу Excel
    val path = "C:/Users/Elizarovy/IdeaProjects/ApacheExample/input.xlsx/"


    val pathTwo = "C:/Users/Elizarovy/IdeaProjects/ApacheExample/result.xlsx/"
//    writeDataToExcel(data, pathTwo)
    interWindow(path, pathTwo)
}

private fun interWindow(path: String, pathTwo: String) {
    SwingUtilities.invokeLater {
        // Создаем новое окно
        val frame = JFrame("Excel")
        frame.defaultCloseOperation = JFrame.EXIT_ON_CLOSE
        frame.setSize(300, 200)

        // Создаем панель
        val panel = JPanel()

        // Создаем первую кнопку
        val button1 = JButton("Open Excel").apply {
            addActionListener {
                val file = File(path)
                Desktop.getDesktop().open(file)
            }
        }

        // Создаем вторую кнопку
        val button2 = JButton("Save and Close").apply {
            addActionListener {
                if (isFileOpen(pathTwo)) {
                    return@addActionListener
                }
                var resultString = ""
                val list = ArrayList<String>()
                val resultList = resultStrokeRow(path, resultString, list)
                val data = getResultMap(resultList)
                writeDataToExcel(data, pathTwo)


                val screen = Screen()

                try {
                    // Ищем окно Excel
                    val excelWindow = screen.find("input.xlsx")

                    if (excelWindow != null) {
                        // Находим кнопку закрытия окна
                        val closeButton = screen.find("close_button.png")

                        if (closeButton != null) {
                            // Кликаем по кнопке закрытия
                            closeButton.click()
                            println("Файл Excel успешно закрыт.")
                        } else {
                            println("Кнопка закрытия не найдена.")
                        }
                    } else {
                        println("Окно Excel не найдено.")
                    }
                } catch (e: Exception) {
                    e.printStackTrace()
                }
                
                val file = File(pathTwo)
                Desktop.getDesktop().open(file)

            }
        }

        // Добавляем кнопки на панель
        panel.add(button1)
        panel.add(button2)

        // Добавляем панель в окно
        frame.contentPane.add(panel)

        // Делаем окно видимым
        frame.isVisible = true
    }
}

fun writeDataToExcel(data: Map<String, Map<Int, Int>>, filePath: String) {
    val size = getMaxKeyFromInnerMaps(data)

    val workbook: Workbook = XSSFWorkbook()
    val sheet = workbook.createSheet("Данные")

    val cellStyle: CellStyle = workbook.createCellStyle().apply {
        alignment = HorizontalAlignment.CENTER
        verticalAlignment = VerticalAlignment.CENTER
    }
    // Создаем заголовки
    val headerRow = sheet.createRow(0)
    headerRow.createCell(0).setCellValue("Ф.И.О.")
    headerRow.getCell(0).cellStyle = cellStyle
    headerRow.createCell(1).setCellValue("Номера упражнений")
    headerRow.getCell(1).cellStyle = cellStyle
    sheet.addMergedRegion(CellRangeAddress(0, 0, 1, size))

    var rowIndex = 1
    var row = sheet.createRow(rowIndex++)
    for (i in 1..size) {
        row.createCell(i).setCellValue(i.toString())
    }
    // Заполнение данными
    for ((name, exercises) in data) {
        row = sheet.createRow(rowIndex++)
        row.createCell(0).setCellValue(name)

        for (i in 1..size) {
            val value = exercises.getOrDefault(i, 0)
            row.createCell(i).setCellValue(value.toDouble())
        }
    }

    // Запись в файл
    FileOutputStream(filePath).use { outputStream ->
        workbook.write(outputStream)
        outputStream.close()
    }

    workbook.close()
}

//Список строк Имя - данные
private fun resultStrokeRow(path: String, resultString: String, list: ArrayList<String>): List<String> {
    var resultStringRow = resultString
    FileInputStream(path).use { fis ->
        // Открываем файл Excel
        val workbook = WorkbookFactory.create(fis)
        // Получаем первый лист
        val sheet = workbook.getSheetAt(0)
        for (rowIndex in 1 until sheet.physicalNumberOfRows) {
            val row = sheet.getRow(rowIndex)
            for (cellIndex in 0 until row.physicalNumberOfCells) {
                if (row.getCell(cellIndex).stringCellValue.isNotEmpty()) {
                    resultStringRow += "${row.getCell(cellIndex).stringCellValue},"
                } else {
                    break
                }
            }
            list.add(resultStringRow)
            resultStringRow = ""
        }
    }
    return list
}

// Преобразование в словарь списка данных по типу Имя = {данные, данные ...}
fun getResultMap(input: List<String>): Map<String, Map<Int, Int>> {
    val result = mutableMapOf<String, MutableMap<Int, Int>>()

    for (entry in input) {
        val parts = entry.split(",").map { it.trim() }
        val name = parts[0]
        val values = parts.drop(1)

        // Получаем или создаем карту для текущего имени
        val valueMap = result.getOrPut(name) { mutableMapOf() }

        for (value in values) {
            if (value == "") continue
            val (key, count) = value.split("-").map { it.toInt() }
            // Суммируем значения по ключу
            valueMap[key] = valueMap.getOrDefault(key, 0) + count
        }
    }
    return result
}

fun getMaxKeyFromInnerMaps(outerMap: Map<String, Map<Int, Int>>): Int {
    return outerMap.values
        .flatMap { it.keys } // Извлекаем все ключи из вложенных карт
        .distinct() // Убираем дубликаты
        .maxOrNull()!! // Находим максимальный ключ
}

fun isFileOpen(filePath: String): Boolean {
    val file = File(filePath)
    return try {
        RandomAccessFile(file, "rw").use {
            false // Если удалось открыть файл для записи, значит он не открыт эксклюзивно другим процессом
        }
    } catch (e: Exception) {
        throw Exception("Файл уже открыт")
        true // Если возникла ошибка при открытии файла для записи, значит он открыт другим процессом
    }
}



