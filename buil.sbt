
organization := "org.realsoftwarematters"

// Set the project name to the string 'My Project'
name := "timesheetaglomerator"

// The := method used in Name and Version is one of two fundamental methods.
// The other method is <<=
// All other initialization methods are implemented in terms of these.
version := "1.0-SNAPSHOT"

scalaVersion := "2.9.0"

// Add a single dependency
libraryDependencies += org.apache.poi" % "poi" % "3.7"


mainClass in (Compile, run) := Some("org.realsoftwarematters.timesheetaglomerator.TimeSheetAglomerator")


