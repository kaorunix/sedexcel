package org.colibrifw.sedexcel

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
//import org.colibrifw.sedexcel.CellManager

object SedExcel {
  def main(args: Array[String]): Unit = {
    println("args="+args.toSeq.toString)
    val (sourceFile, replaceFile, configFile) = args.toList match {
      case (source:String)::(replace:String)::(config:String)::str => (Some(source), Some(replace), Some(config)) 
      case _ => (None, None, None)
    }
    val f = File(configFile.getOrElse(throw new IllegalArgumentException()))
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
  }
  
  def sedCell(cell: Cell, replacestring: Seq[(String, String)]):Unit = {
    //println(f"sedCell $cell $replacestring")
    val datematch = """(\d{4}/\d{1,2}/\d{1,2})""".r
    val nummatch = """(-?\d+.?\d*)""".r
    val strmatch = """(\S)""".r
    replacestring.map (kv => kv._1 match {
      case datematch(k) if cell.getCellType == Cell.CELL_TYPE_NUMERIC && {println("date:" + cell.getNumericCellValue + ":" + kv._2); true} && DateUtil.isCellDateFormatted(cell) && cell.getDateCellValue == (new SimpleDateFormat("yyyy/MM/dd")).parse(k) => println("day="+kv._2);cell.setCellValue(kv._2)
//      case datematch(k) if cell.getCellType == Cell.CELL_TYPE_STRING && {println("date:" + cell.getStringCellValue + ":" + kv._2); true} && DateUtil.isCellDateFormatted(cell) && cell.getStringCellValue == (new SimpleDateFormat("yyyy/MM/dd")).parse(k) => println("day="+kv._2);cell.setCellValue(kv._2)
      case nummatch(k) if cell.getCellType == Cell.CELL_TYPE_NUMERIC && {println("num:" + cell.getNumericCellValue + ":" + kv._2); true} && cell.getNumericCellValue == k.toDouble  => println("num="+kv._2);cell.setCellValue(kv._2)
      case strmatch(k) if cell.getCellType == Cell.CELL_TYPE_STRING && {println("str:" + cell.getStringCellValue + ":" + kv._2); true} && cell.getStringCellValue == k => println("str="+kv._2);cell.setCellValue(kv._2)
      case a => ( println(new CellManager(cell) + "else:"+a) ) }
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
      case e:Exception => throw e
    } finally {
//      is.close() //ここで入力ストリームを閉じる
    }
  }
  
  def writeWorkbook(workbook:Workbook, excelfile: String):Unit = {
    // 編集した内容の書き出し
    val output = File(excelfile)
    workbook.write( output.newOutputStream )
  }
  
  def parseArgs(args: List[String]): Either[Throwable, (String, String, String)] = {
    val (sourceFile, replaceFile, configFile) = args match {
      case (source:String)::(replace:String)::(config:String)::str => (Some(source), Some(replace), Some(config)) 
      case _ => Left(new IllegalArgumentException())
    }
    sourceFile      
    ???
  }
  
   def isFile(path: String):Either[Throwable, Boolean] = {
     try {
       val f = File(path)
       Right(f.isRegularFile && f.size > 0)
     }catch {
       case e:Exception => Left(e)
     }
   }
}