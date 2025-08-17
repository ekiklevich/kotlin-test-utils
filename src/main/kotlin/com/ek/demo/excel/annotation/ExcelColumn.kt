package com.ek.demo.excel.annotation

@Retention(AnnotationRetention.RUNTIME)
@Target(AnnotationTarget.PROPERTY)
annotation class ExcelColumn(
    val header: String = "",
    val format: String = "" // например: "dd.MM.yyyy", "#,##0.00"
)