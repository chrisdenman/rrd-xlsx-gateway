package uk.co.ceilingcat.rrd.gateways.xlsxinputgateway

import arrow.core.Either
import arrow.core.Either.Companion.left
import arrow.core.Either.Companion.right
import arrow.core.flatMap
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import uk.co.ceilingcat.rrd.entities.ServiceDetails
import uk.co.ceilingcat.rrd.entities.ServiceType
import uk.co.ceilingcat.rrd.entities.createServiceDetails
import uk.co.ceilingcat.rrd.usecases.CurrentDate
import uk.co.ceilingcat.rrd.usecases.NextUpcomingInputGateway
import java.io.File
import java.time.LocalDate
import java.time.format.DateTimeFormatter
import java.time.format.DateTimeParseException
import java.util.function.Predicate

data class WorkSheetsSearchDirectory(val file: File)
data class StreetName(val text: String)

sealed class XlsxInputGatewayException : Throwable() {
    object NextUpcomingException : XlsxInputGatewayException()
}

typealias NextUpcomingError = XlsxInputGatewayException.NextUpcomingException

interface XlsxInputGateway : NextUpcomingInputGateway

// @todo check val getters
fun createXlsxInputGateway(
    currentDate: CurrentDate,
    streetName: StreetName,
    workSheetsSearchDirectory: WorkSheetsSearchDirectory
): XlsxInputGateway {

    val recyclingDiscriminator = "recycling"
    val xlsxExtension = "xlsx"
    val xlsxDateFormat = "d MMMM yyyy"

    fun cellContainsStreetName(cell: Cell) = cell.toString() == streetName.text

    val acquireCurrentDate = currentDate.localDate

    fun postfixWithCurrentYear(text: String) = "$text ${acquireCurrentDate.year}"

    fun cellToLocalDate(cell: Cell): Either<NextUpcomingError, LocalDate> = try {
        right(
            LocalDate.parse(
                postfixWithCurrentYear(cell.toString()),
                DateTimeFormatter.ofPattern(xlsxDateFormat)
            )
        )
    } catch (dpe: DateTimeParseException) {
        left(NextUpcomingError)
    }

    fun cellContainsDate(cell: Cell) = cellToLocalDate(cell).isRight()

    fun firstCellThat(sheet: Sheet, predicate: Predicate<Cell>): Either<NextUpcomingError, Cell> {
        val initial: Cell? = null
        val found: Cell? = sheet.fold(initial) { rowAcc, row ->
            row.fold(rowAcc) { cellAcc, cell ->
                if ((cellAcc == null) && predicate.test(cell)) cell else cellAcc
            }
        }
        return if (found == null) {
            left(NextUpcomingError)
        } else {
            right(found)
        }
    }

    fun loadWorkbook(file: File): Either<NextUpcomingError, Workbook> = try {
        right(WorkbookFactory.create(file, null, true))
    } catch (t: Throwable) {
        left(NextUpcomingError)
    }

    fun listWorkBooks(): Either<NextUpcomingError, List<File>> =
        try {
            right(
                workSheetsSearchDirectory
                    .file
                    .listFiles { dir, name ->
                        (name != null) && File(
                            dir,
                            name
                        ).extension.toLowerCase() == xlsxExtension
                    }!!.toList()
            )
        } catch (se: SecurityException) {
            left(NextUpcomingError)
        }

    fun parseServiceType(text: String): ServiceType =
        when {
            text.toLowerCase().contains(recyclingDiscriminator) -> ServiceType.RECYCLING
            else -> ServiceType.REFUSE
        }

    fun getNextUpcomingEntry(sheet: Sheet): Either<NextUpcomingError, ServiceDetails?> {
        val initial: Either<NextUpcomingError, ServiceDetails?> = right(null)
        return firstCellThat(sheet, ::cellContainsStreetName).flatMap { streetNameCell ->
            firstCellThat(sheet, ::cellContainsDate).flatMap { firstDateCell ->
                ((streetNameCell.columnIndex + 1) until (streetNameCell.row.lastCellNum)).fold(initial) { acc, curr ->
                    cellToLocalDate(firstDateCell.row.getCell(curr)).flatMap { cellDate ->
                        createServiceDetails(
                            cellDate,
                            parseServiceType(streetNameCell.row.getCell(curr).toString())
                        ).let { serviceDetails ->
                            acc.flatMap {
                                if ((it == null) || (serviceDetails < it)) right(serviceDetails) else acc
                            }
                        }
                    }
                }
            }
        }
    }

    fun getNextUpcomingEntry(workBook: Workbook): Either<NextUpcomingError, ServiceDetails?> {
        val init: ServiceDetails? = null
        return right(
            workBook
                .sheetIterator()
                .asSequence()
                .toList()
                .fold(init) { acc, sheet ->
                    getNextUpcomingEntry(sheet).fold({ acc }) { serviceDetails ->
                        if (serviceDetails == null) acc else if (acc == null) serviceDetails else if (serviceDetails < acc) serviceDetails else acc
                    }
                }
        )
    }

    return object : XlsxInputGateway {
        override fun nextUpcoming(): Either<NextUpcomingError, ServiceDetails?> {
            val init: ServiceDetails? = null
            return listWorkBooks().map { workBookFiles ->
                workBookFiles.fold(init) { acc, curr ->
                    loadWorkbook(curr).flatMap { workBook ->
                        workBook.use {
                            getNextUpcomingEntry(it).flatMap { serviceDetails ->
                                right(
                                    if (serviceDetails == null) acc else if (acc == null) serviceDetails else if (serviceDetails < acc) serviceDetails else acc
                                )
                            }
                        }
                    }.fold({ acc }, { it })
                }
            }
        }
    }
}
