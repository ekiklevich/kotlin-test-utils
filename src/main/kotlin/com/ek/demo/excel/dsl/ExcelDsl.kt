package com.ek.demo.excel.dsl

import com.ek.demo.excel.ExcelType

fun excel(
    type: ExcelType = ExcelType.XLSX,
    block: ExcelBuilder.() -> Unit
): ExcelBuilder = ExcelBuilder(type).apply(block)