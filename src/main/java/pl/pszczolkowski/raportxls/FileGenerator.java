package pl.pszczolkowski.raportxls;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;

import java.io.*;

class FileGenerator {

    private static final String XLS_EXTENSION = "xlsx";

    private final Workbook workbook;
    private final CellReference cellReference;
    private final String baseDirectory;

    FileGenerator(Workbook workbook, String cellAddress, String baseDirectory) {
        this.workbook = workbook;
        this.cellReference = new CellReference(cellAddress);
        this.baseDirectory = baseDirectory;
    }

    void generate(String fileName) throws IOException {
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(cellReference.getRow());
        if (row == null) {
            row = sheet.createRow(cellReference.getRow());
        }
        Cell cell = row.getCell(cellReference.getCol());
        if (cell == null) {
            cell = row.createCell(cellReference.getCol());
        }
        cell.setCellType(CellType.STRING);

        cell.setCellValue(fileName);

        File file = new File(baseDirectory, fileName + "." + XLS_EXTENSION);
        try (OutputStream outputStream = new FileOutputStream(file)) {
            workbook.write(outputStream);
        }
    }

}
