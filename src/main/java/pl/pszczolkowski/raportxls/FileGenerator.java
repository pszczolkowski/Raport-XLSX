package pl.pszczolkowski.raportxls;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;

import java.io.*;

class FileGenerator {

    private static final String XLS_EXTENSION = "xlsx";

    private final Workbook workbook;
    private final CellReference departmentNameCellReference;
    private final CellReference departmentOwnerCellReference;
    private final String baseDirectory;

    FileGenerator(Workbook workbook, String departmentNameCellAddress, String departmentOwnerCellAddress, String baseDirectory) {
        this.workbook = workbook;
        this.departmentNameCellReference = new CellReference(departmentNameCellAddress);
        this.departmentOwnerCellReference = new CellReference(departmentOwnerCellAddress);
        this.baseDirectory = baseDirectory;
    }

    void generate(Department department) throws IOException {
        Sheet sheet = workbook.getSheetAt(0);

        setCellValue(sheet, departmentNameCellReference, department.getName());
        setCellValue(sheet, departmentOwnerCellReference, department.getOwnerName());

        File file = new File(baseDirectory, department.getName() + "." + XLS_EXTENSION);
        try (OutputStream outputStream = new FileOutputStream(file)) {
            workbook.write(outputStream);
        }
    }

    private void setCellValue(Sheet sheet, CellReference cellReference, String value) {
        Row row = sheet.getRow(cellReference.getRow());
        if (row == null) {
            row = sheet.createRow(cellReference.getRow());
        }

        Cell cell = row.getCell(cellReference.getCol());
        if (cell == null) {
            cell = row.createCell(cellReference.getCol());
        }

        cell.setCellType(CellType.STRING);
        cell.setCellValue(value);
    }

}
