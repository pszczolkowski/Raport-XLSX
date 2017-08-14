package pl.pszczolkowski.raportxls;

import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

class FileNamesReader {

    private FileNamesReader() {}

    static List<String> read(String filePath, String cellAddress) throws InvalidFormatException, IOException {
        try (InputStream inputStream = new FileInputStream(filePath)) {

            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.getSheetAt(0);

            CellReference cellReference = new CellReference(cellAddress);
            int rowIndex = cellReference.getRow();

            List<String> fileNames = new ArrayList<>();
            do {
                Row row = sheet.getRow(rowIndex);
                Cell cell = row.getCell(cellReference.getCol());
                if (cell == null) {
                    break;
                }

                fileNames.add(cell.getStringCellValue());
                rowIndex += 1;
            } while (true);

            return fileNames;
        }
    }

}
