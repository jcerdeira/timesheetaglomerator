package org.realsoftwarematters.timesheetaglomerator

import jxl.{ Workbook, Sheet, Cell, CellType, DateCell, NumberCell, DateFormulaCell, NumberFormulaCell }
import java.io.{ FileInputStream, File }
import java.util.{ Calendar }
import collection.JavaConversions._
import collection.mutable._
import java.lang.reflect.{ Field }
import sun.font.EAttribute
import sun.font.EAttribute

object TimeSheetAglomerator {

  def main(args: Array[String]) = {

    //      println((0 to 20).toList.map( i => extractValues(doc.getSheet("Fevereiro").getCell(1,i))))

    var data: Map[String, List[MonthHours]] = Map()

    for (file <- new File("excels").listFiles) {

      val doc = Workbook.getWorkbook(file)

      println(file.getName())

      if (!doc.getSheetNames().contains("TimeSheet")) {

        doc.getSheetNames().map { sheetName =>
          println("Sheet Name: " + sheetName)
          processSheet(doc.getSheet(sheetName), extractHeadersYear)
        }.foreach { it =>
          val (year, monthHours) = it
          data = addMore(data, year, monthHours)
        }
      } else {
        val (year, monthHours) = processSheet(Workbook.getWorkbook(file).getSheet("TimeSheet"), extractHeadersMonth)
        data = addMore(data, year, monthHours)
      }
    }
    println(data)

  }

  def addMore(data: Map[String, List[MonthHours]], year: String, monthHours: MonthHours): Map[String, List[MonthHours]] = {

    if (!data.contains(year))
      data + (year -> List(monthHours))
    else
      data + (year -> (data(year) ++ List(monthHours)))

  }

  def printLine(sheet: Sheet, pos: Integer) = {

    (0 to 20).toList.map(i => print(" " + extractValues(sheet.getCell(i, pos)) + " "))
    println
  }

  def processSheet(sheet: Sheet, f: Sheet => HeaderSheet): (String, MonthHours) = {
    val headers = f(sheet)
    headers.year -> MonthHours(headers.month, extractMonthValues(headers, extractProjects(sheet), extractActivities(sheet, 6), sheet))
  }

  def extractMonthValues(headers: HeaderSheet, projects: List[IndexLegenda], activities: List[IndexLegenda], sheet: Sheet): List[ProjectHours] = {
    projects map {
      project =>
        ProjectHours(project.value, extractMonthProjectValues(sheet, project.pos, activities))
    }
  }

  def extractMonthProjectValues(sheet: Sheet, nrow: Integer, activities: List[IndexLegenda]): scala.collection.immutable.Map[String, Double] = {
    activities map {
      activity =>
        activity.value -> extractValues(sheet.getCell(activity.pos, nrow)).asInstanceOf[Double]
    } toMap
  }

  def extractHeadersYear(sheet: Sheet): HeaderSheet = {
    val Name = """(.+) / (\d+)""".r
    val Name(month, year) = extractValues(sheet.getCell(9, 1))
    HeaderSheet(translateMonth(month), year)
  }

  def extractHeadersMonth(sheet: Sheet): HeaderSheet = {
    val cal = Calendar.getInstance
    cal.setTime(extractValues(sheet.getCell(16, 1)).asInstanceOf[java.util.Date])
    HeaderSheet(monthName(cal.get(Calendar.MONTH)), cal.get(Calendar.YEAR).toString())
  }

  def monthName(num: Int): String = {
    num match {
      case Calendar.JANUARY => "Janeiro"
      case Calendar.FEBRUARY => "Fevereiro"
      case Calendar.MARCH => "Março"
      case Calendar.APRIL => "Abril"
      case Calendar.MAY => "Maio"
      case Calendar.JUNE => "Junho"
      case Calendar.JULY => "Julho"
      case Calendar.AUGUST => "Agosto"
      case Calendar.SEPTEMBER => "Setembro"
      case Calendar.OCTOBER => "Outubro"
      case Calendar.NOVEMBER => "Novembro"
      case Calendar.DECEMBER => "Dezembro"
    }
  }

  def extractProjects(sheet: Sheet): List[IndexLegenda] = {

    var indexeslegend = List[IndexLegenda]()
    var i = 7
    do {
      val value = extractValues(sheet.getCell(1, i)).asInstanceOf[java.lang.String]
      indexeslegend = indexeslegend ::: List(new IndexLegenda(value = value, pos = i))
      i = i + 1
    } while ((i < 21) && ((extractValues(sheet.getCell(1, i)).asInstanceOf[java.lang.String]).size != 1))

    indexeslegend
  }

  def extractActivities(sheet: Sheet, nrow: Integer): List[IndexLegenda] = {

    var indexeslegend = List[IndexLegenda]()
    for (i <- 2 to 18) {
      val activity = extractValues(sheet.getCell(i, nrow)).asInstanceOf[String]
      if (activity != "") {
        indexeslegend = indexeslegend ::: List(new IndexLegenda(value = translateActivity(activity), pos = i))
      }
    }
    indexeslegend
  }

  def translateMonth(month: String) = month match {
    case "Mar�o" => "Março"
    case _ => month
  }

  def translateActivity(activity: String) = activity match {
    case "Forma��" => "Formação"
    case "Integra��" => "Integração"
    case "Documenta��" => "Documentação"
    case "Reuni�es Externas" => "Reuniões Externas"
    case "Reuni�es Internas" => "Reuniões Internas"
    case "Investiga��" => "Investigação"
    case _ => activity
  }

  def extractValues(cell: Cell): Any = {

    cell.getType() match {
      case CellType.NUMBER => cell.asInstanceOf[NumberCell].getValue()
      case CellType.DATE => cell.asInstanceOf[DateCell].getDate
      case CellType.DATE_FORMULA => cell.asInstanceOf[DateFormulaCell].getDate()
      case CellType.NUMBER_FORMULA => cell.asInstanceOf[NumberFormulaCell].getValue()
      case CellType.STRING_FORMULA => cell.getContents
      case _ => cell.getContents
    }

  }

}

case class YearHous(year: String, months: List[MonthHours]) {}

case class ProjectHoursData(monthProjectHours: List[MonthHours]) {}

case class MonthHours(month: String, projectshours: List[ProjectHours]) {}

case class ProjectHours(projectName: String, activitieshours: scala.collection.immutable.Map[String, Double]) {}

case class HeaderSheet(val month: String, val year: String) {

  override def toString() = {
    getClass().getDeclaredFields().map { field: Field =>
      field.setAccessible(true)
      field.getName() + ": " + field.get(this).toString()
      //println("coisa" + field.getName)
    }.mkString(",")
  }

}

case class IndexLegenda(val value: String, val pos: Integer) {

  override def toString() = {
    getClass().getDeclaredFields().map { field: Field =>
      field.setAccessible(true)
      field.getName() + ": " + field.get(this).toString()
      //println("coisa" + field.getName)
    }.mkString(",")
  }

}
