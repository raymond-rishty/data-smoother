package datasmoothing

import com.google.api.services.sheets.v4.model.BatchUpdateValuesRequest
import com.google.api.services.sheets.v4.model.ValueRange
import java.time.LocalDate
import java.time.LocalDateTime
import java.time.Month
import java.time.format.DateTimeFormatter
import java.time.temporal.ChronoUnit
import java.util.*

object Communion {
  val sheetId = "1_QZL-Sr-WgdxdPhRdNYVlgtTOVs7F5OIVVD8g4EDzJU"
  val inputRange = "Sheet1!A:B"
  val outputRange = "Sheet2!A2:B"
  val ofPattern = DateTimeFormatter.ofPattern("M/d/yyyy")!!

  fun getValues() : List<Pair<LocalDate, Double>> {
    val sheetsService = SheetsQuickstart.getSheetsService()
    val valueRange = sheetsService.spreadsheets().values()
      .get(Communion.sheetId, Communion.inputRange)
      .execute()
    val values =   (valueRange.values as AbstractCollection).toArray()[2] as List<List<String>>
    return values
      .drop(1)
      .filter { it.size == 2 }
      .map{
        LocalDate.parse(it[0], ofPattern).to(Integer.parseInt(it[1]).toDouble()) }
  }
  fun writeSmoothData(map: List<Pair<LocalDateTime, Double?>>) {
    val sheetsService = SheetsQuickstart.getSheetsService()
    val content = ValueRange()
    val formattedContent = map
      .map {
        ChronoUnit.DAYS
          .between(LocalDate.of(1899, Month.DECEMBER, 30).atStartOfDay(), it.first).toInt()
        .to(it.second!!) }
      .map { it.toList() }
    content.setValues(formattedContent)
    content.range = outputRange

    val batchUpdateValuesRequest = BatchUpdateValuesRequest()
    batchUpdateValuesRequest.valueInputOption = "RAW"
    batchUpdateValuesRequest.data = listOf(content)

    sheetsService.spreadsheets().values().batchUpdate(sheetId, batchUpdateValuesRequest).execute()
  }
}

private fun toUnitNumber(date: LocalDate, inputUnit: ChronoUnit): Double = inputUnit
  .between(SmoothData.startDate, date.atStartOfDay())
  .toDouble()

private fun getSmoothData(values1: List<Pair<LocalDate, Double>>): List<Pair<LocalDateTime, Double?>> {
  val values = values1
    .map { toUnitNumber(it.first, ChronoUnit.MONTHS).to(it.second) }
  val (xvalues, smooth) = SmoothData.smoothValues(0.2, values, 1)
  val map = SmoothData.getOutputValues(inputUnit = ChronoUnit.MONTHS,
    outputUnit = ChronoUnit.MONTHS,
    smooth = smooth,
    xvalues = xvalues)
  return map
}

fun main(args: Array<String>) {
  val map = getSmoothData(Communion.getValues())
  Communion.writeSmoothData(map)
}

