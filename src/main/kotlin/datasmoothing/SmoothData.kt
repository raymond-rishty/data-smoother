package datasmoothing

import org.apache.commons.math3.analysis.interpolation.LoessInterpolator
import org.apache.commons.math3.analysis.interpolation.SplineInterpolator
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.time.LocalDate
import java.time.Month
import java.time.ZoneId
import java.time.format.DateTimeFormatter
import java.time.temporal.ChronoUnit
import java.util.*

class SmoothData {
  val startDate = LocalDate.of(2006, Month.JANUARY, 1).atStartOfDay()

  fun readBook(s: String): Workbook = XSSFWorkbook(File(s))

  fun readSheet(workbook: Workbook, sheetName: String) : Sheet = workbook.getSheet(sheetName)

  fun smooth(sheet: Sheet, bandwidth: Double = 0.15, robustnessIters: Int = 5, dataCellIndex: Int, valueCellIndex: Int, outputUnit: ChronoUnit, inputUnit: ChronoUnit) {
    val map = getValues(sheet, dataCellIndex, valueCellIndex, inputUnit)

    val xvalues = map.map { it.first }.toDoubleArray()
    val loessInterpolator = LoessInterpolator(bandwidth, robustnessIters)
    val smooth = loessInterpolator.smooth(xvalues, map.map { it.second }.toDoubleArray())
    val interpolate = SplineInterpolator().interpolate(xvalues, smooth)

    val minutesInUnit = convert(inputUnit, outputUnit)
    var zip = xvalues.zip(smooth).toMap()
    val map1 = xvalues
      .min()!!
      .toInt()
      .rangeTo(xvalues.max()!!.toInt()).step(minutesInUnit).map {
      startDate.plus(it.toLong(), inputUnit).to(
        if (zip.containsKey(it.toDouble())) {
          zip[it.toDouble()]
        } else {
          interpolate.value(it.toDouble())
        }
      )
    }

    println(map1
      .map { "${it.first.format(DateTimeFormatter.ofPattern("MM/dd/yyyy"))}\t ${it.second}" }
      .joinToString(separator = "\n"))
  }

  private fun convert(inputUnit: ChronoUnit, outputUnit: ChronoUnit) =
    outputUnit.duration.toMinutes().toInt() / inputUnit.duration.toMinutes().toInt()

  private fun getValues(sheet: Sheet, dateCellIndex: Int, valueCellIndex: Int, inputUnit: ChronoUnit) : List<Pair<Double, Double>> {
    return sheet
      .map { getDate(it.getCell(dateCellIndex)).to(getAmount(it.getCell(valueCellIndex))) }
      .filter { it.first != null && it.second != null }
      .map { toWeekNumber(it.first!!, inputUnit).to(it.second!!) }
  }

  private fun toWeekNumber(date: Date, inputUnit: ChronoUnit): Double = inputUnit
    .between(startDate, date.toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime())
    .toDouble()


  private fun getDate(cell: Cell?): Date? = if (cell?.cellType == Cell.CELL_TYPE_NUMERIC) {
    cell!!.dateCellValue
  } else {
    null
  }

  private fun getAmount(cell: Cell?): Double? = if (cell?.cellType == Cell.CELL_TYPE_NUMERIC || cell?.cellType == Cell.CELL_TYPE_FORMULA) {
    cell!!.numericCellValue
  } else {
    null
  }

  fun smooth(filename: String,
             sheetName: String,
             bandwidth: Double = 0.15,
             robustnessIters: Int = 5,
             dataCellIndex: Int,
             valueCellIndex: Int, outputUnit: ChronoUnit, inputUnit: ChronoUnit)
    = smooth(readSheet(readBook(filename), sheetName), bandwidth, robustnessIters, dataCellIndex, valueCellIndex, outputUnit, inputUnit)
}

fun main(args : Array<String>) {
  val smoothData = SmoothData()
  smoothData.smooth("""c:\Users\rrishty\Downloads\Attendance.xlsx""", "Sheet1", 0.15, 5, 0, 3, ChronoUnit.WEEKS, ChronoUnit.DAYS)
  //smoothData.smooth("""c:\Users\rrishty\Downloads\Exton_DIAL_DIAL_integration_tests_BuildDurationNetTimeStatistics-7_26_16_9_42_AM.xlsx""", "Exton_DIAL_DIAL_integration_tes", 0.3, 8, 1, 3, ChronoUnit.DAYS, ChronoUnit.MINUTES)

}