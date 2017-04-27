package org.colibrifw.sedexcel

import org.apache.poi.ss.usermodel.Cell

class CellManager(cell:Cell) {
  override def toString:String = {
    cell.getCellType match {
      case Cell.CELL_TYPE_NUMERIC => "type:CELL_TYPE_NUMERIC" + cell.getNumericCellValue + "\n" 
      // 関数（SUMとかIFとか）
      case Cell.CELL_TYPE_FORMULA => "type:CELL_TYPE_FORMULA" + cell.getCellFormula + "\n"
      // 真偽
      case Cell.CELL_TYPE_BOOLEAN => "type:CELL_TYPE_BOOLEAN" + cell.getBooleanCellValue + "\n"
      // 文字列
      case Cell.CELL_TYPE_STRING => "type:CELL_TYPE_STRING" + cell.getRichStringCellValue + "\n"
      // 空
      //case Cell.CELL_TYPE_BLANK => "type:CELL_TYPE_BLANK" + cell.getStringCellValue
      case _ => "" //Not found"
    }
  }
}
