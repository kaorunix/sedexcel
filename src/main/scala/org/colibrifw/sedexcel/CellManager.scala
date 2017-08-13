package org.colibrifw.sedexcel

import org.apache.poi.ss.usermodel.Cell
import java.text.SimpleDateFormat
import java.lang.IllegalStateException
import java.text.Normalizer
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.CellStyle

class CellManager(cell:Cell, wb:Workbook) {
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
//    cell.getCellType match {
//      case Cell.CELL_TYPE_NUMERIC => {
        if (value matches """[\d]{4}/[\d]{2}/[\d]{2}""") Right {
          cell.setCellValue(new SimpleDateFormat("yyyy/MM/dd").parse(value))
          val cellStyle = wb.createCellStyle()
          val creationHelper = wb.getCreationHelper()
          cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy/mm/dd"))
          cell.setCellStyle(cellStyle)
        }
        else if (value matches """[\d]{1,2}月[\d]{2}/[\d]{2}""") Right(cell.setCellValue(new SimpleDateFormat("MM月dd日").parse(value)))
        else if (value matches """[\d]{1,}""") Right(cell.setCellValue(value.toInt))
        else Right(cell.setCellValue(Normalizer.normalize(value, Normalizer.Form.NFC))) //Left(new IllegalStateException("CELL_TYPE_FORMULA"))
        
/*      }
      // 関数（SUMとかIFとか）
      case Cell.CELL_TYPE_FORMULA => println(f"CELL_TYPE_FORMULA $value");Left(new IllegalStateException("CELL_TYPE_FORMULA"))
      // 真偽
      case Cell.CELL_TYPE_BOOLEAN => println(f"CELL_TYPE_BOOLEAN $value");Left(new IllegalStateException("CELL_TYPE_BOOLEAN"))
      // 文字列
      case Cell.CELL_TYPE_STRING => println(f"CELL_TYPE_STRING $value");Right(cell.setCellValue(value))
      // 空
      //case Cell.CELL_TYPE_BLANK => "type:CELL_TYPE_BLANK" + cell.getStringCellValue
      case _ => println(f"CELL TYPE NOT FOUND $value");Left(new IllegalStateException("Could not find CELL_TYPE"))
    }*/
  }
}
