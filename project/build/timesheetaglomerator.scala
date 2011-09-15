import sbt._
import de.element34.sbteclipsify._

class timesheetaglomerator(info: ProjectInfo) extends DefaultProject(info) with Eclipsify {

 override def libraryDependencies = Seq(
    "org.apache.poi" % "poi" % "3.7"
  ) ++ super.libraryDependencies

 val poi = "org.apache.poi" % "poi" % "3.7"

  override def mainClass = Some("org.realsoftwarematters.timesheetaglomerator.TimeSheetAglomerator")
}
