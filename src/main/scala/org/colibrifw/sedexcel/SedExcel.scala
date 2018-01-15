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
    replaceCell(sourceFile.get, replaceFile.get, rep.toSeq)
  }
  
  /**
   * key, valueに設定を保持する
   */
  def splitKV(s: String):Option[(String, String)] = {
    val pattarnMatch = """([^=]\S+)=(.+$)""".r // position=value
    s match {
      case pattarnMatch(k,v) => Some(k,v)
      case _ => None
    }
  }
  
  /**
   * @param fromexcelfile Excelテンプレートファイルパス
   * @param toexcelfil 出力Excelファイルパス
   * @param positionString 置き換えExcel座標と文字列
   */
  def replaceCell(fromexcelfile: String, toexcelfile: String, positionString: Seq[(String,String)]):Unit = {
    println(f"replaceCell positionString=$positionString")
    val wb = getWorkbook(fromexcelfile)
    for (
         (position, value) <- positionString;
         (sheet, x, y) <- expositon2positon(position);
         cell <- Some(wb.getSheet(sheet).getRow(y).getCell(x))
    ) yield {
      val c = new CellManager(cell, wb)
      c.setValue(value)
    }
    wb.getCreationHelper().createFormulaEvaluator().evaluateAll()
    writeWorkbook(wb, toexcelfile) match {
      case Right(_) => println("Process Success")
      case Left(e) => println(f"Failed to save file ${toexcelfile} :${e.getMessage}, ${e.printStackTrace()}")
    }
  }
  
  /**
   * Excelのセル位置から数値の座標へ変換
   */
  def expositon2positon(position: String): Option[(String, Int, Int)] = {
    val patternXY = "^([^:]*):([a-zA-Z]{1,3})([0-9]{1,10})".r
    position match {
      case patternXY(a: String, b: String, c: String) if (c.toInt > 0) => Some((a, a2i(b), c.toInt - 1)) 
      case _ => None
    }
  }
  
  /**
   * Excelのセル位置を表すアルファベットを数値に置き換える
   */
  def a2i(ascii: String): Int = {
    val unit = 'Z'.toInt - 'A'.toInt + 1
    val asciilist = ascii.toUpperCase  // 大文字にそろえる
         .toList       // Charに変換
         .map(a => a.toInt - 64) // ASCIIコードから数字に変換
    asciilist     
         .foldLeft(0)((b,c) => b * unit + c) - 1// 前桁
  }

  def getWorkbook(excelfile: String):Workbook = {
    val sourcef = File(excelfile)
    val is = sourcef.newInputStream
    try {
      val wb = WorkbookFactory.create( is ) // メモリ上に展開
      is.close
      wb
    } catch {
      case e:Exception => throw e
    } finally {
      is.close() //ここで入力ストリームを閉じる
    }
  }

  /**
   * Workbookオブジェクトを出力する
   * @param workbook 編集後Workbookオブジェクト
   * @param execfile ファイルパス
   */
  def writeWorkbook(workbook:Workbook, excelfile: String):Either[Exception, Unit] = {
    // 編集した内容の書き出し
    val output = File(excelfile)
    println("writeWorkbook")
    try {
      workbook.write( output.newOutputStream )
      Right(Unit)
    } catch {
      case e: java.io.IOException =>  Left(e) 
    }
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