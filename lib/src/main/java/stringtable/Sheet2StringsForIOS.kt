package stringtable

import org.apache.poi.ss.usermodel.Sheet
import stringtable.xls.SheetNavigator
import java.io.File
import java.io.IOException
import java.lang.StringBuilder
import java.nio.charset.Charset
import kotlin.NoSuchElementException

object Sheet2StringsForIOS {
    fun convert(sheet: Sheet?, pathRes: File) {
        val nav = SheetNavigator(sheet)
        val columnStringId = findColumnId(nav)

        var i = 1
        while (true) {
            val column = i++
            try {
                val languageCode = nav.getCell(0, column)
                if ("values" !in languageCode) continue
                val iosLanguageFileName = "${languageCode.replace("values-", "")}.lproj"
                val filename = Path.combine(pathRes.path, iosLanguageFileName, "Localizable.strings")
                createStringsXml(filename, nav, columnStringId, column)
            } catch (e: NoSuchElementException) {
                break
            }
        }
    }

    private fun findColumnId(nav: SheetNavigator): Int {
        var result = 0
        while (true) {
            result++
            try {
                if (isIdColumn(nav.getCell(0, result).toLowerCase()))
                    return result
            } catch (e: NoSuchElementException) {
                return 0
            }
        }
    }

    private fun isIdColumn(columnHeaderName: String): Boolean {
        if (columnHeaderName.startsWith("id "))
            return true
        if (columnHeaderName.endsWith(" id"))
            return true
        if (columnHeaderName in listOf("id", "ids", "identification", "identifications"))
            return true

        return false
    }

    private fun createStringsXml(filename: String, nav: SheetNavigator, columnStringId: Int, col: Int) {

        val doc = StringBuilder()
        doc.append("// generator link: https://github.com/DNights/android-string-table\n")
        doc.append("// original link: https://github.com/jobtools/android-string-table\n")
        doc.append("// 자동 생성된 파일입니다. 이 xml 파일을 직접 수정하지 마세요!\n")
        doc.append("// This file is auto generated. DO NOT EDIT THIS XML FILE!\n")

        var row = 1
        while (true) {
            var id: String
            var value: String
            try {
                id = nav.getCell(row, columnStringId)
                id = id.trim { it <= ' ' }
                if (id.isEmpty()) {
                    row++
                    continue
                }

                // 주석처리
                if (id.startsWith("<") ||
                        id.startsWith("/") ||
                        id.startsWith("#")) {
                    if (id.startsWith("<!--")) {
                        id = id.replace("<!--", "").replace("-->".toRegex(), "")
                    }
                    var commentString = id
                    val description = nav.getCell(row, 1)
                    if (!description.isEmpty()) {
                        commentString = "$commentString / $description "
                    }

                    doc.append("//$commentString\n")

                    row++
                    continue
                }
                value = try {
                    nav.getCell(row, col)
                } catch (e: NoSuchElementException) {
                    ""
                }
                if (value.isEmpty()) {
                    row++
                    continue
                }
            } catch (e: NoSuchElementException) {
                break
            }

            if (id.contains("[]")) {
                val stringArray = getStringArrayItem(id, value)
                doc.append(stringArray)
            } else {
              doc.append("\"$id\" = \"$value\";\n")
            }
            row++
        }

        try {
            val file = File(filename)

            if (!file.parentFile.isDirectory)
                file.parentFile.mkdirs()
            else if (file.exists())
                file.delete()

            file.writeText(doc.toString(), Charset.forName("UTF-8"))

        } catch (e: IOException) {
            e.printStackTrace()
        }
    }

    private fun getStringArrayItem(id: String, value: String): String {
        val str = StringBuilder()
        val bagicId = id.replace("[]", "")

        var separator = "\n"
        if (value.contains("_x000a_")) separator = "_x000a_"
        val items = value.split(separator.toRegex()).toTypedArray()
        for ((index, item) in items.withIndex()) {
            str.append("\"$bagicId[$index]\" = \"$item\";\n")
        }
        return str.toString()
    }
}