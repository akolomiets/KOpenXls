package com.npk.documents

import com.npk.GeneratedDocumentPath
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.IndexedColors
import org.junit.jupiter.api.Test
import java.io.OutputStream
import java.nio.file.Files.deleteIfExists
import java.nio.file.Files.newOutputStream
import java.nio.file.Path
import java.nio.file.StandardOpenOption.CREATE

internal class KOpenXlsTest {

    @GeneratedDocumentPath
    private lateinit var generatedPdfPath: Path

    @Test
    fun `generate a showcase`(): Unit = generatedPdfPath.resolve("showcase.xlsx").let { documentPath ->
        deleteIfExists(documentPath)
        newOutputStream(documentPath, CREATE).use { output -> buildShowcaseWorkBook(output) }
    }

    private fun buildShowcaseWorkBook(outputStream: OutputStream) = buildXSSFWorkbook(outputStream) {
        sheet {
            columnWidth["A"] = 10.0
            columnWidth["B"] = 10.0
            columnWidth["C"] = columnWidth["A"] + columnWidth["B"]

            row {
                val cell1 = cell(1) {
                }
                val cell2 = cell(1)
                formula("sum($cell1, $cell2)")
            }

            skip(2)

            table(listOf("Q1", "Q2", "Q3")) {
                style {
                    name = "TableStyleLight13"
                    showColumnStripes = true
                    showRowStripes = false
                }

                row {
                    val c1 = cell(1)
                    val c2 = cell(2)
                    formula("sum($c1, $c2)")
                }
                row {
                    val c1 = cell(3)
                    val c2 = cell(4)
                    formula("sum($c1, $c2)")
                }
                row {
                    val c1 = cell(5)
                    val c2 = cell(6)
                    formula("sum($c1, $c2)")
                }
            }

            row {
                cell("Done!") {
                    cellStyle {
                        fillPattern = FillPatternType.SOLID_FOREGROUND
                        fillBackgroundColor = IndexedColors.TAN.index
                    }
                }
            }
        }
    }


    @Test
    fun `generate the table`(): Unit = generatedPdfPath.resolve("table.xlsx").let { documentPath ->
        deleteIfExists(documentPath)
        newOutputStream(documentPath, CREATE).use { output ->
            buildXSSFWorkbook(output) {
                sheet("Table") {
                    columnWidths(10.0)
                    table("No", "Column1", "Column2", "Column3", "Formula1") {
                        style {
                            name = "TableStyleLight13"
                            showColumnStripes = true
                            showRowStripes = false
                        }

                        repeat(1000) {
                            row {
                                cell(it + 1)
                            }
                        }
                    }
                }
            }
        }
    }

}
