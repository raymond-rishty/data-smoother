package datasmoothing

import com.google.api.client.repackaged.com.google.common.base.CharMatcher
import com.google.api.services.sheets.v4.model.BatchUpdateValuesRequest
import com.google.api.services.sheets.v4.model.ValueRange
import com.google.common.primitives.Doubles
import java.time.LocalDate
import java.time.LocalDateTime
import java.time.Month
import java.time.format.DateTimeFormatter
import java.time.temporal.ChronoUnit
import java.util.*

object Attendance {
  val sheetId = "1e58BYaD4-D79m4_36v3cogp1IEuOThqvcaW5xFIhJvI"
  val inputRanges =
  listOf(
    "Offering".to("Sheet1!A:D"),
    "Adults".to("Sheet1!A:E"),
    "Derivative".to("Sheet7!A:K")
  ).toMap()
  val outputRanges = listOf(
    "Offering".to("Sheet7!A2:B"),
    "Adults".to("Sheet7!D2:E"),
    "Derivative".to("Sheet7!M3:N")
  ).toMap()
  val ofPattern = DateTimeFormatter.ofPattern("M/d/yyyy")!!

  fun getValues(dimension: String): List<Pair<LocalDate, Double>> {

    val sheetsService = SheetsQuickstart.getSheetsService()
    val valueRange = sheetsService.spreadsheets().values()
      .get(Attendance.sheetId, inputRanges[dimension])
      .execute()
    val values =   (valueRange.values as AbstractCollection).toArray()[2] as List<List<String>>
    val charMatcher = CharMatcher.`is`('$').or(CharMatcher.`is`(','))
    val expectedSize = values.map { it.size }.max()
    return values
      .drop(1)
      .filter { it.size == expectedSize }
      .map{
        LocalDate.parse(it.first(), ofPattern).to(Doubles.tryParse(charMatcher.removeFrom(it.last()))!!)
      }
  }
  fun writeSmoothData(dimension: String, map: List<Pair<LocalDateTime, Double?>>) {
    val sheetsService = SheetsQuickstart.getSheetsService()
    val content = ValueRange()
    val formattedContent = map
      .map {
        ChronoUnit.DAYS
          .between(LocalDate.of(1899, Month.DECEMBER, 30).atStartOfDay(), it.first).toInt()
        .to(it.second!!) }
      .map { it.toList() }
    content.setValues(formattedContent)
    content.range = outputRanges[dimension]

    val batchUpdateValuesRequest = BatchUpdateValuesRequest()
    batchUpdateValuesRequest.valueInputOption = "RAW"
    batchUpdateValuesRequest.data = listOf(content)

    sheetsService.spreadsheets().values().batchUpdate(sheetId, batchUpdateValuesRequest).execute()
  }
}

private fun toUnitNumber(date: LocalDate, inputUnit: ChronoUnit): Double = inputUnit
  .between(SmoothData.startDate, date.atStartOfDay())
  .toDouble()

private fun getSmoothData(values1: List<Pair<LocalDate, Double>>, bandwidth: Double, robustIters: Int): List<Pair<LocalDateTime, Double?>> {
  val outputUnit = ChronoUnit.WEEKS
  val values = values1
    .map { toUnitNumber(it.first, outputUnit).to(it.second) }
  val (xvalues, smooth) = SmoothData.smoothValues(bandwidth, values, robustIters)
  val map = SmoothData.getOutputValues(inputUnit = outputUnit,
    outputUnit = outputUnit,
    smooth = smooth,
    xvalues = xvalues)
  return map
}

fun main(args: Array<String>) {
  Attendance.writeSmoothData("Offering", getSmoothData(Attendance.getValues("Offering"), 0.15, 1))
  Attendance.writeSmoothData("Adults", getSmoothData(Attendance.getValues("Adults"), 0.3, 2))
  Attendance.writeSmoothData("Derivative", getSmoothData(Attendance.getValues("Derivative"), 0.05, 0))
}

