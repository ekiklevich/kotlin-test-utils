package com.ek.demo.excel.dsl

import com.ek.demo.excel.ExcelType
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.ByteArrayOutputStream
import java.io.OutputStream
import java.nio.file.Files
import java.nio.file.Path

class ExcelBuilder(type: ExcelType) {
    private var index = 1
    private val workbook: Workbook = when (type) {
        ExcelType.XLS -> HSSFWorkbook()
        ExcelType.XLSX -> XSSFWorkbook()
    }

    fun sheet(
        name: String = "Sheet-${index++}",
        block: SheetBuilder.() -> Unit
    ) = apply {
        val sh = workbook.createSheet(name)
        SheetBuilder(workbook, sh).apply(block)
    }

    fun writeTo(path: String) {
        writeTo(Path.of(path))
    }

    fun writeTo(path: Path) {
        Files.newOutputStream(path).use { os -> writeTo(os) }
    }

    fun writeTo(output: OutputStream) {
        workbook.use { wb -> wb.write(output) }
    }

    fun toByteArray(): ByteArray {
        val bos = ByteArrayOutputStream()
        workbook.write(bos)
        return bos.toByteArray()
    }
}