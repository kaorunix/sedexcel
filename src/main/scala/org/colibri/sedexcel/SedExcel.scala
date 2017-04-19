package org.colibri.sedexcel

import better.files._
import java.io.FileInputStream
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import scala.collection.JavaConversions._
import java.text.SimpleDateFormat
import java.lang.IllegalArgumentException

object SedExcel {
  val configFilePath = "config.txt"

  def main(args: Array[String]): Unit = {
    println("args="+args.toSeq.toString)
    val (sourceFile, replaceFile, configFile) = args.toList match {
      case (source:String)::(replace:String)::(config:String)::str => (Some(source), Some(replace), Some(config)) 
      case _ => (None, None, None)
    }
    val f = File(configFile.getOrElse(throw new IllegalArgumentException))
//    f.lines.foreach { line => splitKV(line).map(kv => kv._1, kv._2) }
    val rep = for (
        line <- f.lines;
        kv <- splitKV(line)
      ) yield (kv)
    sedexcel(sourceFile.get, replaceFile.get, rep.toSeq)
  }
  
  def splitKV(s: String):Option[(String, String)] = {
    val pattarnMatch = """([^=]\S+)=(\S+)""".r
    s match {
      case pattarnMatch(k,v) => Some(k,v)
      case _ => None
    }
  }
  
  def sedexcel(fromexcelfile:String, toexcelfile:String, replacestrings:Seq[(String, String)]) = {
    val wb = getWorkbook(fromexcelfile)  
    for (
        sheet <- (0 to wb.getNumberOfSheets() - 1).map(i => wb.getSheetAt(i));
        row <- sheet.iterator();
        cell <- row.iterator()
    ) sedCell(cell, replacestrings)
    writeWorkbook(wb, toexcelfile)

    // 編集したいシート、列、セルを指定
/*    val s:Sheet =wb.getSheetAt(0)
    val r:Row = s.getRow(1)
    val c:Cell = r.getCell(1)
// この場合B2セルに「あいう」をセット
c.setCellValue( "あいう" );
c.setCellType( Cell.CELL_TYPE_STRING );*/

  }
  
  def sedCell(cell: Cell, replacestring: Seq[(String, String)]):Unit = {
    val datematch = """(\d{4}/\d{1,2}/\d{1,2})""".r
    val nummatch = """(-?\d+.?\d*)""".r
    val strmatch = """(\S)""".r
    replacestring.map (kv => kv._1 match {
      case datematch(k) if cell.getCellType == Cell.CELL_TYPE_NUMERIC && DateUtil.isCellDateFormatted(cell) && cell.getDateCellValue == (new SimpleDateFormat("yyyy/MM/dd")).parse(k) => println("date="+kv._2.toString);cell.setCellValue(kv._2)
      case nummatch(k) if cell.getCellType == Cell.CELL_TYPE_NUMERIC && cell.getNumericCellValue == k.toDouble => println("num="+kv._2.toString);cell.setCellValue(kv._2)
      case strmatch(k) if cell.getCellType == Cell.CELL_TYPE_STRING && cell.getRichStringCellValue.getString == k => println("str="+kv._2.toString);cell.setCellValue(kv._2)
      case _ => () }
    )
  }

  def getWorkbook(excelfile: String):Workbook = {
    try {
      val sourcef = File(excelfile)
      val is = sourcef.newInputStream
      val wb = WorkbookFactory.create( is ) // メモリ上に展開
      is.close
      wb
    } catch {
      case e:Exception => ??? 
    } finally {
//      is.close() //ここで入力ストリームを閉じる
    }
  }
  
  def writeWorkbook(workbook:Workbook, excelfile: String):Unit = {
    println("writeWorkbook")
    // 編集した内容の書き出し
    val output = File(excelfile)
    workbook.write( output.newOutputStream )
  }
}