package org.colibrifw.sedexcel

import org.apache.poi.ss.usermodel.Cell
import java.text.SimpleDateFormat
import java.lang.IllegalStateException

class CellManager(cell:Cell) {
  override def toString:String = {
    "style" ++ cell.getCellStyle.getDataFormatString ++ "::" ++{
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
    }}
  }
  
  def setValue(value: String): Either[IllegalStateException, Unit] = {
    val yyyymmdd = """[\d]{4}/[\d]{2}/[\d]{2}""".r
    val tsukihi = """[\d]{1,2}月[\d]{2}/[\d]{2}""".r
/*    val sdf: SimpleDateFormat = value matche {
    case """[\d]{4}/[\d]{2}/[\d]{2}""" => 
      new SimpleDateFormat("yyyyMMdd HH:mm:ss");
    }*/
    cell.getCellType match {
      case Cell.CELL_TYPE_NUMERIC => {
        if (value matches """[\d]{4}/[\d]{2}/[\d]{2}""") Right(cell.setCellValue(new SimpleDateFormat("yyyy/MM/dd").parse(value)))
        else if (value matches """[\d]{1,2}月[\d]{2}/[\d]{2}""") Right(cell.setCellValue(new SimpleDateFormat("yyyy/MM/dd").parse(value)))
        else if (value matches """[\d]{1,}""") Right(cell.setCellValue(value.toInt))
        else Left(new IllegalStateException("CELL_TYPE_FORMULA"))
      }
      // 関数（SUMとかIFとか）
      case Cell.CELL_TYPE_FORMULA => Left(new IllegalStateException("CELL_TYPE_FORMULA"))
      // 真偽
      case Cell.CELL_TYPE_BOOLEAN => Left(new IllegalStateException("CELL_TYPE_BOOLEAN"))
      // 文字列
      case Cell.CELL_TYPE_STRING => Right(cell.setCellValue(value))
      // 空
      //case Cell.CELL_TYPE_BLANK => "type:CELL_TYPE_BLANK" + cell.getStringCellValue
      case _ => Left(new IllegalStateException("Could not find CELL_TYPE"))
    }
  }
}
