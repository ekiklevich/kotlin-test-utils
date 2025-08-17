package com.ek.demo.excel.dsl

import com.ek.demo.excel.annotation.ExcelColumn
import org.apache.poi.ss.usermodel.*
import java.time.*
import java.util.*
import kotlin.reflect.KClass
import kotlin.reflect.KProperty1
import kotlin.reflect.full.findAnnotation
import kotlin.reflect.full.memberProperties
import kotlin.reflect.full.primaryConstructor

class SheetBuilder(
    workbook: Workbook,
    private val sheet: Sheet
) {
    private val defaultDateFormat = "yyyy-MM-dd"
    private val defaultDateTimeFormat = "yyyy-MM-dd HH:mm"
    private val defaultZonedDateTimeFormat = "yyyy-MM-dd HH:mm z"
    private val styles = DefaultStyles(workbook)

    inline fun <reified T : Any> from(
        items: List<T>,
        includeHeader: Boolean = true
    ) {
        from(T::class, items, includeHeader)
    }

    fun <T : Any> from(
        type: KClass<T>,
        items: List<T>,
        includeHeader: Boolean
    ) {
        val cols = resolveColumns(type)

        var rowIndex = 0
        if (includeHeader) {
            val headerRow = sheet.createRow(rowIndex++)
            cols.forEachIndexed { i, col ->
                val cell = headerRow.createCell(i)
                cell.setCellValue(col.header)
                cell.cellStyle = styles.header
            }
        }

        for (item in items) {
            val row = sheet.createRow(rowIndex++)
            cols.forEachIndexed { i, col ->
                val cell = row.createCell(i)
                val value = col.property.get(item)
                writeCell(cell, value, col.format)
            }
        }

        sheet.createFreezePane(0, if (includeHeader) 1 else 0)
        cols.indices.forEach { sheet.autoSizeColumn(it) }
    }

    private fun writeCell(cell: Cell, value: Any?, format: String?) {
        val style = format?.takeIf { it.isNotBlank() }?.let { styles.custom(it) }

        when (value) {
            null -> cell.setBlank()
            is Number -> {
                cell.setCellValue(value.toDouble())
                style?.let { cell.cellStyle = it }
            }
            is Boolean -> cell.setCellValue(value)
            is LocalDate -> {
                val date = Date.from(value.atStartOfDay(ZoneId.systemDefault()).toInstant())
                cell.setCellValue(date)
                cell.cellStyle = style ?: styles.date(defaultDateFormat)
            }
            is LocalDateTime -> {
                val date = Date.from(value.atZone(ZoneId.systemDefault()).toInstant())
                cell.setCellValue(date)
                cell.cellStyle = style ?: styles.date(defaultDateTimeFormat)
            }
            is ZonedDateTime -> {
                val date = Date.from(value.toInstant())
                cell.setCellValue(date)
                cell.cellStyle = style ?: styles.date(defaultZonedDateTimeFormat)
            }
            is Instant -> {
                val date = Date.from(value.atZone(ZoneId.systemDefault()).toInstant())
                cell.setCellValue(date)
                cell.cellStyle = style ?: styles.date(defaultDateTimeFormat)
            }
            is Date -> {
                cell.setCellValue(value)
                cell.cellStyle = style ?: styles.date(defaultDateTimeFormat)
            }
            else -> cell.setCellValue(value.toString())
        }
    }
}

private data class ResolvedColumn<T : Any>(
    val property: KProperty1<T, *>,
    val header: String,
    val format: String?
)

private fun <T : Any> resolveColumns(kClass: KClass<T>): List<ResolvedColumn<T>> {
    val ctorOrder = kClass.primaryConstructor
        ?.parameters
        ?.mapIndexedNotNull { idx, p -> p.name?.let { it to idx } }
        ?.toMap()
        ?: emptyMap()

    val props = kClass.memberProperties

    return props
        .filter { it.name in ctorOrder }
        .sortedBy { ctorOrder[it.name] }
        .map { prop ->
            val ann = prop.findAnnotation<ExcelColumn>()
            val header = ann?.header?.takeIf { it.isNotBlank() } ?: prop.name
            val format = ann?.format?.takeIf { it.isNotBlank() }
            ResolvedColumn(prop, header, format)
        }
}

private class DefaultStyles(val workbook: Workbook) {
    val header: CellStyle = workbook.createCellStyle().apply {
        val font = workbook.createFont().apply { bold = true }
        setFont(font)
        verticalAlignment = VerticalAlignment.CENTER
    }

    fun date(fmt: String): CellStyle = custom(fmt)

    fun custom(fmt: String): CellStyle =
        workbook.createCellStyle().apply {
            val helper = workbook.creationHelper
            dataFormat = helper.createDataFormat().getFormat(fmt)
        }
}