package pl.sda;

import org.apache.commons.lang3.tuple.ImmutablePair;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

public class ExcelReader {

    private static final String FILE_PATH = "c:\\budzet_kowalskich.xls";

    public Pair<List, List> read() throws IOException, InvalidFormatException {

        try (InputStream inp = new FileInputStream(FILE_PATH)) {
            Workbook wb = WorkbookFactory.create(inp);
            Sheet sheet = wb.getSheetAt(0);
            for (int i = 1; i < sheet.getLastRowNum(); i++) {

                Row row = sheet.getRow(i);
                BigDecimal income = getCellValue(row, 1);
                BigDecimal outcome = getCellValue(row, 3);

                System.out.println("Dochód: " + income);
                System.out.println("Wydatek: " + outcome);
            }
        }
        return new ImmutablePair<>(new ArrayList<>(), new ArrayList<>());
    }

    private BigDecimal getCellValue(Row row, int columnNumber) {
        Cell cell = row.getCell(columnNumber);
        return !cell.toString().isEmpty() ?
                new BigDecimal(cell.toString()) : BigDecimal.ZERO;
    }

    public static void main(String[] args) throws IOException, InvalidFormatException {
        new ExcelReader().read();
    }
}
