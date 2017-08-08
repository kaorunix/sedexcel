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
  val datePrefix = "DATE"
//  val numberPrefix = "NUM"
  val pricePrefix = "PRICE"
  val stringPrefix = "STR"
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
  
  /**
   * key, valueに設定を保持する
   */
  def splitKV(s: String):Option[(String, String)] = {
    val pattarnMatch = """([^=]\S+)=(\S+)""".r
    s match {
      case pattarnMatch(k,v) if (Seq(datePrefix, pricePrefix, stringPrefix).exists { k.startsWith(_)}) => Some(k,v)
      case _ => None
    }
  }
  
  /**
   * @param positionString 置き換え文字列
   * @param fromexcelfile Excelテンプレートファイルパス
   * @param toexcelfil 出力Excelファイルパス
   */
  def replaceCell(positionString: Seq[(String,String)] ,fromexcelfile: String, toexcelfile: String):Unit = {
    val wb = getWorkbook(fromexcelfile)
    for (
         (position, value) <- positionString;
         (sheet, x, y) <- expositon2positon(position);
         cell <- Some(wb.getSheet(sheet).getRow(y).getCell(x))
         
    ) yield 
    writeWorkbook(wb, toexcelfile)
  }
  
  /**
   * Excelのセル位置から数値の座標へ変換
   */
  def expositon2positon(position: String): Option[(String, Int, Int)] = {
    val patternXY = "^([^:]*):([a-zA-Z]{1,3})([0-9]){1,10}".r
    position match {
      case patternXY(a: String, b: String, c: String) => Some((a, a2n(b), c.toInt)) 
      case _ => None
    }
  }
  
  /**
   * Excelのセル位置を表すアルファベットを数値に置き換える
   */
  def a2n(ascii: String): Int = {
    val unit = 'Z'.toInt - 'A'.toInt
    ascii.toUpperCase  // 大文字にそろえる
         .toList       // Charに変換
         .map(a => a.toInt - 65) // ASCIIコードから数字に変換
         .foldLeft(0)((b,c) => b * unit + c) // 前桁
  }
  
  /**
   * @param c 変更
   */
  def setCellValueByType(c: Cell, value: String) = {
    case class CellValue(p: String)
    val cell = new CellManager(c)
    cell.setValue(value)      
    }

    /**
   * 文字列を置き換える
   */
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
    println(new CellManager(cell))
    Seq(datePrefix, pricePrefix, stringPrefix).map {
      prefix => {
        replacestring.filter(
            cell.getCellType == Cell.CELL_TYPE_STRING
            && cell.getStringCellValue.startsWith(datePrefix) 
            && _._1.startsWith(datePrefix)
            ).map {
              kv => if (kv._1 == cell.getRichStringCellValue) { println("kv._1=" + kv._1);cell.setCellValue(kv._2) }
            }
      }
    }
/*    replacestring.map (kv => kv._1 match {
      case datematch(k) if cell.getCellType == Cell.CELL_TYPE_NUMERIC && {println("date:" + cell.getNumericCellValue + ":" + kv._2); true} && DateUtil.isCellDateFormatted(cell) && cell.getDateCellValue == (new SimpleDateFormat("yyyy/MM/dd")).parse(k) => println("day="+kv._2);cell.setCellValue(kv._2)
//      case datematch(k) if cell.getCellType == Cell.CELL_TYPE_STRING && {println("date:" + cell.getStringCellValue + ":" + kv._2); true} && DateUtil.isCellDateFormatted(cell) && cell.getStringCellValue == (new SimpleDateFormat("yyyy/MM/dd")).parse(k) => println("day="+kv._2);cell.setCellValue(kv._2)
      case nummatch(k) if cell.getCellType == Cell.CELL_TYPE_NUMERIC && {println("num:" + cell.getNumericCellValue + ":" + kv._2); true} && cell.getNumericCellValue == k.toDouble  => println("num="+kv._2);cell.setCellValue(kv._2)
      case strmatch(k) if cell.getCellType == Cell.CELL_TYPE_STRING && {println("str:" + cell.getStringCellValue + ":" + kv._2); true} && cell.getStringCellValue == k => println("str="+kv._2);cell.setCellValue(kv._2)
      case a => ( println(new CellManager(cell) + "else:"+a) ) }
    )
    def compCellDate(c:Cell, value:String):Boolean = {
      DateUtil.isCellDateFormatted(cell) && cell.getDateCellValue == (new SimpleDateFormat("yyyy/MM/dd")).parse(value)      
    }*/
    /*def setCellValueByType(prefix: String, cell: Cell, value: String) = {
      case class CellValue(p: String)
      prefix match {
        case "DATE" => print("DATE " + value);cell.setCellType(Cell.CELL_TYPE_NUMERIC);cell.setCellValue(value)
        case "PRICE" => print("PRICE " + value);cell.setCellType(Cell.CELL_TYPE_NUMERIC);cell.setCellValue(value) 
        //case CellValue(pricePrefix) => cell.setCellType(Cell.CELL_TYPE_NUMERIC);cell.setCellValue(value)
        case "STR" => print("STR" + value);cell.setCellType(Cell.CELL_TYPE_STRING);cell.setCellValue(value)
      }
      
    }*/
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