package com.npk.documents

import org.apache.poi.ss.util.CellAddress
import org.apache.poi.ss.util.CellUtil
import org.apache.poi.ss.util.CellUtil.DATA_FORMAT
import org.apache.poi.xssf.usermodel.*
import java.io.OutputStream
import java.time.Instant
import java.time.LocalDate
import java.time.LocalDateTime
import java.util.*


/**
 *
 */
fun buildXSSFWorkbook(outputStream: OutputStream, block: ExcelWorkbookBuilder.() -> Unit) =
    XSSFWorkbook().use { workbook ->
        ExcelWorkbookBuilder(workbook).apply(block)
        workbook.write(outputStream)
    }


interface IndexedAccessOperator<T> {

    operator fun get(index: Int): T

    operator fun set(index: Int, value: T)

}

@DslMarker
internal annotation class ExcelBuilderDslMarker

@ExcelBuilderDslMarker
class ExcelWorkbookBuilder(val workbook: XSSFWorkbook) {

    /**
     *
     */
    fun sheet(name: String = "", block: ExcelSheetBuilder.() -> Unit) {
        val sheetName = name.ifEmpty { "Sheet${workbook.numberOfSheets + 1}" }
        val sheet = workbook.createSheet(sheetName)
        ExcelSheetBuilder(workbook, sheet).apply(block)
    }

}

@ExcelBuilderDslMarker
class ExcelSheetBuilder(val workbook: XSSFWorkbook, val sheet: XSSFSheet) {

    private var skipRows = 1

    /**
     *
     */
    val columnWidth = object : IndexedAccessOperator<Double> {
        override fun get(index: Int): Double = sheet.getColumnWidth(index) / 256.0
        override fun set(index: Int, value: Double) = sheet.setColumnWidth(index, (value * 256).toInt())
    }

    /**
     *
     */
    fun columnWidths(vararg widths: Double) {
        widths.forEachIndexed { index, width ->
            if (width >= 0) {
                columnWidth[index] = width
            }
        }
    }

    /**
     *
     */
    fun columnStyle(index: Int, block: XSSFCellStyle.() -> Unit) {
        sheet.setDefaultColumnStyle(index, workbook.createCellStyle().apply(block))
    }

    /**
     *
     */
    fun row(block: ExcelRowBuilder.() -> Unit) {
        val row = sheet.createRow(sheet.lastRowNum + skipRows).also { skipRows = 1 }
        ExcelRowBuilder(workbook, row).apply(block)
    }

    /**
     *
     */
    fun skip(n: Int = 1) {
        skipRows = (n + 1)
    }

}

@ExcelBuilderDslMarker
class ExcelRowBuilder(val workbook: XSSFWorkbook, val row: XSSFRow) {

    private companion object {
        const val DEFAULT_DATA_FORMAT_CODE = 14
    }

    private var currentCellNum = row.lastCellNum.toInt()

    /**
     *
     */
    fun cell(value: Any?, block: XSSFCell.() -> Unit = {}): CellAddress =
        row.createCell(++currentCellNum)
            .let { cell ->
                when (value) {
                    null -> cell.setBlank()
                    is Boolean -> cell.setCellValue(value)
                    is Number -> cell.setCellValue(value.toDouble())
                    is Date -> cell.setDateFormatProperty().setCellValue(value)
                    is Calendar -> cell.setDateFormatProperty().setCellValue(value)
                    is LocalDate -> cell.setDateFormatProperty().setCellValue(value)
                    is LocalDateTime -> cell.setDateFormatProperty().setCellValue(value)
                    is Instant -> cell.setDateFormatProperty().setCellValue(Date.from(value))
                    else -> cell.setCellValue(value.toString())
                }
                CellAddress(cell.apply(block))
            }

    /**
     *
     */
    fun formula(content: String, block: XSSFCell.() -> Unit = {}): CellAddress =
        row.createCell(++currentCellNum)
            .let { cell ->
                cell.cellFormula = content
                CellAddress(cell.apply(block))
            }

    /**
     *
     */
    fun XSSFCell.cellStyle(block: XSSFCellStyle.() -> Unit) {
        cellStyle = workbook.createCellStyle().apply(block)
    }

    private fun XSSFCell.setDateFormatProperty(): XSSFCell =
        this.also { CellUtil.setCellStyleProperty(it, DATA_FORMAT, DEFAULT_DATA_FORMAT_CODE) }

}
