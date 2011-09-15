package org.realsoftwarematters.timesheetaglomerator

import org.apache.poi.hssf.usermodel.{ HSSFWorkbook }
import org.apache.poi.poifs.filesystem.POIFSFileSystem
import org.apache.poi.hssf.extractor.ExcelExtractor
import org.apache.poi.ss.usermodel.{ Cell, Row, Sheet, DateUtil }

import java.io.{ FileInputStream, File }
import java.util.{ Calendar }

import collection.JavaConversions._
import collection.mutable._

import java.lang.reflect.{ Field }

object TimeSheetAglomerator {

  def main(args: Array[String]) = {

    for (file <- new File("excels").listFiles) {
      val sheet = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(file))).getSheet("TimeSheet")
      println(processSheet(sheet))
    }

  }

  def printRow(row: Row) = {

    row.cellIterator foreach { cell => print("index: %s value: [%s] " format (cell.getColumnIndex(), extractValues(cell))) }

  }

  def processSheet(sheet: Sheet): (String, MonthHours) = {
    val headers = extractHeaders(sheet)
    headers.year -> MonthHours(headers.month, extractMonthValues(headers, extractProjects(sheet), extractActivities(sheet.getRow(6)), sheet))
  }

  def extractMonthValues(headers: HeaderSheet, projects: List[IndexLegenda], activities: List[IndexLegenda], sheet: Sheet): List[ProjectHours] = {
    projects map {
      project =>
        ProjectHours(project.value, extractMonthProjectValues(sheet.getRow(project.pos), activities))
    }
  }

  def extractMonthProjectValues(row: Row, activities: List[IndexLegenda]): scala.collection.immutable.Map[String, Double] = {
    activities map (
      activity =>
        activity.value -> extractValues(row.getCell(activity.pos)).asInstanceOf[Double]) toMap
  }

  def extractHeaders(sheet: Sheet): HeaderSheet = {
    val cal = Calendar.getInstance
    cal.setTime(extractValues(sheet.getRow(1).getCell(16)).asInstanceOf[java.util.Date])
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
      case Calendar.OCTOBER => "Outubro"
      case Calendar.NOVEMBER => "Novembro"
      case Calendar.DECEMBER => "Dezembro"
    }
  }

  def extractProjects(sheet: Sheet): List[IndexLegenda] = {

    var indexeslegend = List[IndexLegenda]()
    var i = 7; var value = "test"
    do {
      value = extractValues(sheet.getRow(i).getCell(1)).asInstanceOf[java.lang.String]
      indexeslegend = indexeslegend ::: List(new IndexLegenda(value = value, pos = i))
      i = i + 1
    } while ((i < 21) && ((extractValues(sheet.getRow(i).getCell(1)).asInstanceOf[java.lang.String]).size != 1))

    indexeslegend
  }

  def extractActivities(row: Row): List[IndexLegenda] = {

    var indexeslegend = List[IndexLegenda]()
    for (i <- 2 to 18) {
      if (extractValues(row.getCell(i)).asInstanceOf[String] != "") indexeslegend = indexeslegend ::: List(new IndexLegenda(value = row.getCell(i).toString, pos = i))
    }
    indexeslegend
  }

  def extractValues(cell: Cell): Any = {

    cell.getCellType match {
      case Cell.CELL_TYPE_NUMERIC =>
        if (DateUtil.isCellDateFormatted(cell)) { cell.getDateCellValue }
        else { cell.getNumericCellValue }
      case Cell.CELL_TYPE_BLANK => ""
      case _ => cell.getStringCellValue
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

case class IndexLegenda(val value: String, val pos: Int) {

  override def toString() = {
    getClass().getDeclaredFields().map { field: Field =>
      field.setAccessible(true)
      field.getName() + ": " + field.get(this).toString()
      //println("coisa" + field.getName)
    }.mkString(",")
  }

}
