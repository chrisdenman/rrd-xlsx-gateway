package uk.co.ceilingcat.rrd.gateways.xlsxinputgateway

import org.junit.jupiter.api.Assertions
import org.junit.jupiter.api.Test
import org.junit.jupiter.api.TestInstance
import org.junit.jupiter.api.TestInstance.Lifecycle.PER_CLASS
import uk.co.ceilingcat.rrd.usecases.createCurrentDate
import java.io.File

@TestInstance(PER_CLASS)
internal class XlsxInputGatewayTests {

    @Test
    fun `That createXlsxInputGateway() returns instances with ane values`() {
        val currentDate = createCurrentDate()
        val streetName = StreetName("The Mall")
        val worksheetsSearchDirectory = WorkSheetsSearchDirectory(File("resources"))
        Assertions.assertNotNull(createXlsxInputGateway(currentDate, streetName, worksheetsSearchDirectory))
    }
}
