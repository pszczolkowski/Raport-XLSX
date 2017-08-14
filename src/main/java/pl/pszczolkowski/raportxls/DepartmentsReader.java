package pl.pszczolkowski.raportxls;

import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

class DepartmentsReader {

    private DepartmentsReader() {}

    static List<Department> read(String filePath, String cellAddress) throws InvalidFormatException, IOException {
        try (InputStream inputStream = new FileInputStream(filePath)) {

            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.getSheetAt(0);

            CellReference cellReference = new CellReference(cellAddress);
            int rowIndex = cellReference.getRow();

            List<Department> departments = new ArrayList<>();
            do {
                Row row = sheet.getRow(rowIndex);
                if (row == null) {
                    break;
                }

                Cell nameCell = row.getCell(cellReference.getCol());
                if (nameCell == null) {
                    break;
                }
                Cell ownerCell = row.getCell(cellReference.getCol() + 1);
                if (ownerCell == null) {
                    throw new RuntimeException("no matching owner cell for cell "
                            + CellReference.convertNumToColString(cellReference.getCol()) + rowIndex);
                } else if (ownerCell.getCellTypeEnum() != CellType.STRING) {
                    throw new RuntimeException("cell " + CellReference.convertNumToColString(cellReference.getCol())
                            + rowIndex + " with owner name is not of text type");
                }

                Department department = new Department(nameCell.getStringCellValue(), ownerCell.getStringCellValue());
                departments.add(department);

                rowIndex += 1;
            } while (true);

            return departments;
        }
    }

}
