package org.colibrifw.sedexcel

import org.scalatest._
import better.files._

class SedExcelSpec extends FlatSpec with Matchers {
  "splitKV" should "get parametors" in {
    SedExcel.splitKV("""sheetname:C1=123""") should be (Some(("sheetname:C1", "123")))
    SedExcel.splitKV("""sheetname:C1=abc""") should be (Some(("sheetname:C1", "abc")))
    SedExcel.splitKV("""sheetname:C1=2017/04/27""") should be (Some(("sheetname:C1", "2017/04/27")))
    SedExcel.splitKV("""シート名:ZZZ645000=文字列可能 スペースの後ろも取得""") should be (Some(("シート名:ZZZ645000", "文字列可能 スペースの後ろも取得")))
    SedExcel.splitKV("""abc123""") should be (None)
    SedExcel.splitKV("""=12345""") should be (None)
    SedExcel.splitKV("""#コメント行のつもり""") should be (None)
  }

  "splitKV" should "get parametors from config file" in {
    val f = File("test/configfile")
    f.lines.foreach { 
      line => SedExcel.splitKV(line).map{
        kv => kv match {
          case ("STRabc", "123") => true
          case ("PRICEabc", "12345") => true
          case ("DATEabc", "2017/04/27") => true
          case _ => fail
        } 
      }
    }
  }

  "expositon2positon" should "calculate position" in {
    SedExcel.expositon2positon("sheet:A1") should be (Some("sheet", 0, 0))
    SedExcel.expositon2positon("sheet:b2") should be (Some("sheet", 1, 1))
    SedExcel.expositon2positon("シート:Z100") should be (Some("シート", 25, 99))
    SedExcel.expositon2positon("パピプペポ:ZZZ65400") should be (Some("パピプペポ", 15625, 65399))
  }

  "a2n" should "change Excel ABC position to number" in {
    SedExcel.a2i("a") should be (0)
    SedExcel.a2i("B") should be (1)
    SedExcel.a2i("c") should be (2)
    SedExcel.a2i("D") should be (3)
    SedExcel.a2i("e") should be (4)
    SedExcel.a2i("F") should be (5)
    SedExcel.a2i("g") should be (6)
    SedExcel.a2i("H") should be (7)
    SedExcel.a2i("i") should be (8)
    SedExcel.a2i("J") should be (9)
    SedExcel.a2i("k") should be (10)
    SedExcel.a2i("L") should be (11)
    SedExcel.a2i("m") should be (12)
    SedExcel.a2i("N") should be (13)
    SedExcel.a2i("o") should be (14)
    SedExcel.a2i("P") should be (15)
    SedExcel.a2i("q") should be (16)
    SedExcel.a2i("R") should be (17)
    SedExcel.a2i("s") should be (18)
    SedExcel.a2i("T") should be (19)
    SedExcel.a2i("u") should be (20)
    SedExcel.a2i("V") should be (21)
    SedExcel.a2i("w") should be (22)
    SedExcel.a2i("X") should be (23)
    SedExcel.a2i("y") should be (24)
    SedExcel.a2i("Z") should be (25)
    SedExcel.a2i("AA") should be (26) // 26 * 1 + 1
    SedExcel.a2i("ZA") should be (676)
    SedExcel.a2i("ZZ") should be (701)
    SedExcel.a2i("ZZA") should be (18226)
  }
  
  "getWorkbook" should "open excel file" in {
    val wk = SedExcel.getWorkbook("test/テンプレート.xlsx")
    println("WORKBOOK:" + wk.toString)
    true should === (true)
  }
}
