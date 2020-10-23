package stringtable

import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.security.InvalidParameterException

object StringTableGenerator {

    @Throws(Exception::class)
    @JvmStatic
    fun main(args: Array<String>) {
        if (args.size < 3) {
            throw InvalidParameterException("<platform (AOS,IOS)> <source xlsx> <res path> <sheet name>")
        }
        println("Generate string tables.")
        println("\tplatform: " + args[0])
        println("\tsource: " + args[1])
        println("\tres: " + args[2])
        val platform = args[0]
        val source = File(args[1])
        val pathRes = File(args[2])
        val inputStream = FileInputStream(source)
        val workbook = XSSFWorkbook(inputStream)
        val targetSheetName = if (args.size > 3) args[3] else null

        when(platform){
            "AOS" -> { Sheet2Strings.convert(getTargetSheet(targetSheetName, workbook), pathRes)}
            "IOS" -> { Sheet2StringsForIOS.convert(getTargetSheet(targetSheetName, workbook), pathRes)}
            else -> throw InvalidParameterException("Can you enter platform is AOS or IOS")
        }

        println("Completed.")
        inputStream.close()
    }

    private fun getTargetSheet(targetSheetName: String?, workbook: XSSFWorkbook): XSSFSheet {
        if (targetSheetName == null)
            return workbook.getSheetAt(0)

        return workbook.getSheet(targetSheetName).also { result ->
            assert(result != null) { "Not found target sheet: $targetSheetName" }
        }
    }
}