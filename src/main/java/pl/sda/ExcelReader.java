package pl.sda;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;

public class ExcelReader {

    private static final String FILE_PATH = "c:\\budzet_kowalskich.xls";

    public void read() throws IOException, InvalidFormatException {

        try (InputStream inp = new FileInputStream(FILE_PATH)) {
            Workbook wb = WorkbookFactory.create(inp);
            Sheet sheet = wb.getSheetAt(0);
            for (int i = 1; i < sheet.getLastRowNum(); i++) {

                Row row = sheet.getRow(i);
                BigDecimal income = getCellValue(row, 1);
                BigDecimal outcome = getCellValue(row, 3);
            }
        }
    }

    private BigDecimal getCellValue(Row row, int rowNumber) {
        Cell cell = row.getCell(rowNumber);
        return !cell.toString().isEmpty() ?
                new BigDecimal(cell.toString()) : BigDecimal.ZERO;
    }

    public static void main(String[] args) throws IOException, InvalidFormatException {
        new ExcelReader().read();
    }
}
