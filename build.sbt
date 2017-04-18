name := """sedexcel"""

version := "1.0"

scalaVersion := "2.11.8"

// Change this to another test framework if you prefer
libraryDependencies ++= Seq(
	"org.scalatest" %% "scalatest" % "2.2.4" % "test",
	"org.apache.poi" % "poi" % "3.9",
	"org.apache.poi" % "poi-ooxml" % "3.9",
	"org.apache.poi" % "ooxml-schemas" % "1.3",
	"com.github.pathikrit" %% "better-files" % "2.14.0")

// Uncomment to use Akka
//libraryDependencies += "com.typesafe.akka" %% "akka-actor" % "2.3.11"

