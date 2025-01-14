package com.npk.documents

import org.apache.poi.ss.util.CellReference
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableStyleInfo
import java.util.concurrent.atomic.AtomicReference

/**
 *
 */
operator fun <T> IndexedAccessOperator<T>.get(ref: String): T = get(CellReference.convertColStringToIndex(ref))
operator fun <T> IndexedAccessOperator<T>.set(ref: String, value: T) = set(CellReference.convertColStringToIndex(ref), value)

/**
 *
 */
fun ExcelRowBuilder.cell(value: Any?, style: XSSFCellStyle, block: XSSFCell.() -> Unit = {}) = cell(value) {
    cellStyle = style
    apply(block)
}

/**
 *
 */
fun ExcelSheetBuilder.table(columnNames: List<String>, block: ExcelTableBuilder.() -> Unit) {
    require(columnNames.isNotEmpty()) { "columnNames must not be empty" }
    row {
        columnNames.forEach { name -> cell(name) }
    }
    val topRowNum = sheet.lastRowNum
    val tableStyleBlockRef = AtomicReference<(CTTableStyleInfo) -> Unit>()

    ExcelTableBuilder(workbook, sheet, tableStyleBlockRef).apply(block)
    if (topRowNum == sheet.lastRowNum) {
        row { /* Add empty row */ }
    }

    val areaRef = workbook.creationHelper.createAreaReference(CellReference(topRowNum, 0), CellReference(sheet.lastRowNum, columnNames.size - 1))
    sheet.createTable(areaRef).let { table ->
        with(table.ctTable) {
            columnNames.forEachIndexed { index, _ -> tableColumns.getTableColumnArray(index).id = index + 1L }

            addNewTableStyleInfo().let { style ->
                tableStyleBlockRef.plain
                    ?.let { it(style) }
                    ?: style.apply {
                        name = "TableStyleMedium9"
                        showColumnStripes = true
                        showRowStripes = false
                    }
            }

            addNewAutoFilter().ref = table.area.formatAsString()
        }
    }
}

/**
 *
 */
fun ExcelSheetBuilder.table(vararg columnNames: String, block: ExcelTableBuilder.() -> Unit) = table(columnNames.toList(), block)

@ExcelBuilderDslMarker
class ExcelTableBuilder(val workbook: XSSFWorkbook, val sheet: XSSFSheet, private val tableStyleBlockRef: AtomicReference<(CTTableStyleInfo) -> Unit>) {

    /**
     *
     */
    fun style(block: CTTableStyleInfo.() -> Unit) {
        tableStyleBlockRef.plain = block
    }

    /**
     *
     */
    fun row(block: ExcelRowBuilder.() -> Unit) {
        val row = sheet.createRow(sheet.lastRowNum + 1)
        ExcelRowBuilder(workbook, row).apply(block)
    }

}
