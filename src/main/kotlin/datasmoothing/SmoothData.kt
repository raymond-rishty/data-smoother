package datasmoothing

import org.apache.commons.math3.analysis.interpolation.LoessInterpolator
import org.apache.commons.math3.analysis.interpolation.SplineInterpolator
import org.apache.commons.math3.analysis.polynomials.PolynomialSplineFunction
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.time.LocalDate
import java.time.LocalDateTime
import java.time.Month
import java.time.ZoneId
import java.time.format.DateTimeFormatter
import java.time.temporal.ChronoUnit
import java.util.*

class SmoothData {
  val startDate: LocalDateTime = LocalDate.of(2006, Month.JANUARY, 1).atStartOfDay()

  fun readBook(s: String): Workbook = XSSFWorkbook(File(s))

  fun readSheet(workbook: Workbook, sheetName: String) : Sheet = workbook.getSheet(sheetName)

  fun smooth(sheet: Sheet, bandwidth: Double = 0.15, robustnessIters: Int = 5, dataCellIndex: Int, valueCellIndex: Int, outputUnit: ChronoUnit, inputUnit: ChronoUnit, mapFunction: (inputUnit: ChronoUnit, interpolate: PolynomialSplineFunction, outputUnit: ChronoUnit, xvalues: DoubleArray, zip: Map<Double, Double>) -> List<Pair<LocalDateTime, Double?>> = {inputUnit, interpolate, outputUnit, xvalues, zip -> mapInterpolating(inputUnit, interpolate, outputUnit, xvalues, zip)}) {
    val (xvalues, smooth) = getSmoothValues(bandwidth, dataCellIndex, inputUnit, robustnessIters, sheet, valueCellIndex)
    val interpolate = SplineInterpolator().interpolate(xvalues, smooth)

    val zip = xvalues.zip(smooth).toMap()
    val map = mapFunction(inputUnit, interpolate, outputUnit, xvalues, zip)

    println(map
      .map { "${it.first.format(DateTimeFormatter.ofPattern("MM/dd/yyyy"))}\t ${it.second}" }
      .joinToString(separator = "\n"))
  }

  @Suppress("UNUSED_PARAMETER")
  fun mapWithoutInterpolating(inputUnit: ChronoUnit, interpolate: PolynomialSplineFunction, outputUnit: ChronoUnit, xvalues: DoubleArray, zip: Map<Double, Double>): List<Pair<LocalDateTime, Double>> {
    return zip.map {
      startDate.plus(it.key.toLong(), inputUnit).to(it.value)
    }
  }

  fun mapInterpolating(inputUnit: ChronoUnit, interpolate: PolynomialSplineFunction, outputUnit: ChronoUnit, xvalues: DoubleArray, zip: Map<Double, Double>): List<Pair<LocalDateTime, Double?>> {
    return IntRange(xvalues.min()!!.toInt(), xvalues.max()!!.toInt())
      .step(convert(inputUnit, outputUnit))
      .map {
        startDate.plus(it.toLong(), inputUnit).to(
          if (zip.containsKey(it.toDouble())) {
            zip[it.toDouble()]
          } else {
            interpolate.value(it.toDouble())
          }
        )
      }
  }

  private fun getSmoothValues(bandwidth: Double, dataCellIndex: Int, inputUnit: ChronoUnit, robustnessIters: Int, sheet: Sheet, valueCellIndex: Int): Pair<DoubleArray, DoubleArray> {
    val map = getValues(sheet, dataCellIndex, valueCellIndex, inputUnit)

    val xvalues = map.map { it.first }.toDoubleArray()
    val loessInterpolator = LoessInterpolator(bandwidth, robustnessIters)
    val smooth = loessInterpolator.smooth(xvalues, map.map { it.second }.toDoubleArray())
    return Pair(xvalues, smooth)
  }

  private fun convert(inputUnit: ChronoUnit, outputUnit: ChronoUnit) =
    (outputUnit.duration.toMinutes().toLong() / inputUnit.duration.toMinutes().toLong()).toInt()

  private fun getValues(sheet: Sheet, dateCellIndex: Int, valueCellIndex: Int, inputUnit: ChronoUnit) : List<Pair<Double, Double>> {
    return sheet
      .map { getDate(it.getCell(dateCellIndex)).to(getAmount(it.getCell(valueCellIndex))) }
      .filter { it.first != null && it.second != null }
      .map { toUnitNumber(it.first!!, inputUnit).to(it.second!!) }
  }

  private fun toUnitNumber(date: Date, inputUnit: ChronoUnit): Double = inputUnit
    .between(startDate, date.toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime())
    .toDouble()


  private fun getDate(cell: Cell?): Date? = if (cell?.cellType == Cell.CELL_TYPE_NUMERIC || cell?.cellType == Cell.CELL_TYPE_FORMULA) {
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
             valueCellIndex: Int,
             outputUnit: ChronoUnit,
             inputUnit: ChronoUnit,
             mapFunction: (ChronoUnit, PolynomialSplineFunction, ChronoUnit, DoubleArray, Map<Double, Double>) -> List<Pair<LocalDateTime, Double?>>  = {inputUnit, interpolate, outputUnit, xvalues, zip -> mapInterpolating(inputUnit, interpolate, outputUnit, xvalues, zip)}
  ) = smooth(
    sheet = readSheet(readBook(filename), sheetName),
    bandwidth = bandwidth,
    robustnessIters = robustnessIters,
    dataCellIndex = dataCellIndex,
    valueCellIndex = valueCellIndex,
    outputUnit = outputUnit,
    inputUnit = inputUnit,
    mapFunction = mapFunction
  )
}

fun main(args: Array<String>) {
  val smoothData = SmoothData()
  // offering
/*  smoothData.smooth(
    filename = """c:\Users\rrishty\Downloads\Attendance (5).xlsx""",
    sheetName = "Sheet1",
    bandwidth = 0.15,
    robustnessIters = 3,
    dataCellIndex = 0,
    valueCellIndex = 3,
    outputUnit = ChronoUnit.WEEKS,
    inputUnit = ChronoUnit.DAYS)*/
  // derivative
  smoothData.smooth("""c:\Users\rrishty\Downloads\Attendance (6).xlsx""", "Sheet7", 0.05, 0, 0, 10, ChronoUnit.WEEKS, ChronoUnit.DAYS)
  // adult attendance
/*
  smoothData.smooth(
    filename = """c:\Users\rrishty\Downloads\Attendance (5).xlsx""",
    sheetName = "Sheet1",
    bandwidth = 0.3,
    robustnessIters = 0,
    dataCellIndex = 0,
    valueCellIndex = 4,
    outputUnit = ChronoUnit.WEEKS,
    inputUnit = ChronoUnit.DAYS)
*/
  //smoothData.smooth("""c:\Users\rrishty\Downloads\Exton_DIAL_DIAL_integration_tests_BuildDurationNetTimeStatistics-7_26_16_9_42_AM.xlsx""", "Exton_DIAL_DIAL_integration_tes", 0.3, 8, 1, 3, ChronoUnit.DAYS, ChronoUnit.MINUTES)
  //smoothData.smooth("""c:\Users\rrishty\Documents\realcleanpolitics-ge.xlsx""", "Sheet1", 0.15, 5, 9, 4, ChronoUnit.DAYS, ChronoUnit.HOURS);
  //smoothData.smooth("""c:\Users\rrishty\Documents\realcleanpolitics-ge.xlsx""", "Sheet1", 0.15, 5, 9, 5, ChronoUnit.DAYS, ChronoUnit.HOURS);
/*
  smoothData.smooth(filename = """c:\Users\rrishty\Documents\realclearpolitics-ge.xlsx""",
    sheetName = "Sheet1",
    bandwidth = 0.35,
    robustnessIters = 4,
    dataCellIndex = 10,
    valueCellIndex = 5,
    outputUnit = ChronoUnit.DAYS,
    inputUnit = ChronoUnit.HOURS)
*//*
  smoothData.smooth(filename = """c:\Users\rrishty\Documents\trends.xlsx""",
    sheetName = "trends",
    bandwidth = 0.2,
    robustnessIters = 1,
    dataCellIndex = 0,
    valueCellIndex = 3,
    outputUnit = ChronoUnit.MONTHS,
    inputUnit = ChronoUnit.MONTHS)*/

/*
  smoothData.smooth(filename = """c:\Users\rrishty\Downloads\MEMB.xls (2).xlsx""",
    sheetName = "Sheet5",
    bandwidth = 0.15,
    robustnessIters = 0,
    dataCellIndex = 4,
    valueCellIndex = 6,
    outputUnit = ChronoUnit.MONTHS,
    inputUnit = ChronoUnit.DAYS, mapFunction = { inputUnit, interpolate, outputUnit, xvalues, zip -> smoothData.mapWithoutInterpolating(inputUnit, interpolate, outputUnit, xvalues, zip) })
*/

/*
  smoothData.smooth(filename = """c:\Users\rrishty\Downloads\MEMB.xls (2).xlsx""",
    sheetName = "Sheet5",
    bandwidth = 0.15,
    robustnessIters = 0,
    dataCellIndex = 4,
    valueCellIndex = 5,
    outputUnit = ChronoUnit.MONTHS,
    inputUnit = ChronoUnit.DAYS, mapFunction = { inputUnit, interpolate, outputUnit, xvalues, zip -> smoothData.mapWithoutInterpolating(inputUnit, interpolate, outputUnit, xvalues, zip) })
*/

}