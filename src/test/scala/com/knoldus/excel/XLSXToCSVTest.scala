package com.knoldus.excel

import java.io.File

import org.scalatest.{BeforeAndAfterAll, FunSuite}

import scala.io.Source

class XLSXToCSVTest extends FunSuite with BeforeAndAfterAll {

  val reader = new XLSXToCSV()

  override def afterAll() = {
    val testFile = new File("./target/testresult.csv")
    if(testFile.exists()) testFile.delete()
  }

  test("must convert xlsx file into csv") {
    val inputString = Source.fromFile("./src/test/resources/testresult.csv").getLines().reduce((a,b)=> a+b).trim
    reader.processOneSheet("./src/test/resources/testfile.xlsx","./target/testresult.csv",",")
    val outputString = Source.fromFile("./target/testresult.csv").getLines().reduce((a,b) => a+b).trim
    inputString === outputString
  }

}
