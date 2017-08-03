package org.colibrifw.sedexcel

import org.scalatest._
import better.files._

class SedExcelSpec extends FlatSpec with Matchers {
  "splitKV" should "get parametors" in {
    SedExcel.splitKV("""STRabc=123""") should be (Some(("STRabc", "123")))
    SedExcel.splitKV("""PRICEabc=12345""") should be (Some(("PRICEabc", "12345")))
    SedExcel.splitKV("""DATEabc=2017/04/27""") should be (Some(("DATEabc", "2017/04/27")))
    SedExcel.splitKV("""abc=123""") should be (None)
    SedExcel.splitKV("""99999=12345""") should be (None)
    SedExcel.splitKV("""2999/1/1=2017/04/27""") should be (None)
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

  "getWorkbook" should "open excel file" in {
    val wk = SedExcel.getWorkbook("test/テンプレート.xlsx")
    println("WORKBOOK:" + wk.toString)
    true should === (true)
  }
/*  "Hello" should "have tests" in {
    SedExcel.getWorkbook
    true should === (true)
  }*/
}
