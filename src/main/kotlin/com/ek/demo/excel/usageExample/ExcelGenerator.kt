package com.ek.demo.excel.usageExample

import com.ek.demo.excel.ExcelType
import com.ek.demo.excel.dsl.excel
import com.ek.demo.model.Compensation
import org.springframework.stereotype.Service
import java.math.BigDecimal
import java.time.LocalDate
import java.time.format.DateTimeFormatter

@Service
class ExcelGenerator {

    fun generateXlsx() {
        val compensationsXlsx = listOf(
            Compensation("111", null, null),
            Compensation("222", "222.456", LocalDate.now())
        )

        excel(ExcelType.XLSX) {
            sheet { from(compensationsXlsx) }
        }.writeTo("compensations.xlsx")
    }

    fun generateXls() {
        val compensationsXls = listOf(
            Compensation("334", BigDecimal("334.33"), LocalDate.now().format(DateTimeFormatter.ISO_LOCAL_DATE)),
            Compensation("445", BigDecimal("445.44").toPlainString(), LocalDate.now())
        )

        excel(ExcelType.XLS) {
            sheet { from(compensationsXls) }
        }.writeTo("compensations.xls")
    }
}