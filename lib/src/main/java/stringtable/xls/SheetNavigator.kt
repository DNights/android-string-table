package stringtable.xls

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.xssf.usermodel.XSSFCell
import java.util.*

class SheetNavigator(var sheet: Sheet?) {
    @Throws(NoSuchElementException::class)
    fun getCell(row: Int, col: Int): String {
        val rowData = sheet?.getRow(row)
                ?: throw NoSuchElementException("row is null (" + row + ", " + col + ") on " + sheet?.sheetName)
        val cell = rowData.getCell(col)
                ?: throw NoSuchElementException("Cell is null (" + row + ", " + col + ") on " + sheet?.sheetName)
        return getCellString(cell)
    }

    companion object {
        private fun getCellString(cell: Cell): String {
            val cellType = cell.cellType
            when (cellType) {
                XSSFCell.CELL_TYPE_NUMERIC -> return "" + cell.numericCellValue
                XSSFCell.CELL_TYPE_STRING -> return cell.stringCellValue.replace("_x000a_", "\n")
            }
            return ""
        }
    }
}