name := "spark-excel"

version := "1.0"

val spark = "org.apache.spark" % "spark-core_2.10" % "1.6.0" 
val apacchePOI = "org.apache.poi" % "poi" % "3.13"
val apachePOIXML = "org.apache.poi" % "poi-ooxml" % "3.13"
val scalaTest = "org.scalatest" % "scalatest_2.10" % "3.0.0-M15"
val apacheXCERS = "xerces" % "xercesImpl" % "2.11.0"


lazy val commonSettings = Seq(
  organization := "com.knoldus.spark-excel",
  version := "0.1.0",
  scalaVersion := "2.10.5"
)

lazy val root = (project in file(".")).
  settings(commonSettings: _*).
  settings(
    name := "spark-excel",
    libraryDependencies ++= Seq(spark,apacchePOI,apachePOIXML,scalaTest,apacheXCERS)
  )
