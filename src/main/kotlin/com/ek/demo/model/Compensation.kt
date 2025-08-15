package com.ek.demo.model

import com.ek.demo.excel.annotation.ExcelColumn

data class Compensation(
    val mpContractId: Any?,
    val amount: Any?,
    @ExcelColumn(format = "dd.MM.yyyy")
    val date: Any?
)
