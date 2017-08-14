package pl.pszczolkowski.raportxls;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

public class Main {

    public static void main(String[] args) {
        if (args.length != 4) {
            System.out.println("invalid number of arguments");
            System.out.println("required arguments are: <template_file_path> <destination_cell> <filenames_file_path> <source_cell>");
        }

        String templateFilePath = args[0];
        String destinationCellAddress = args[1];
        String fileNamesFilePath = args[2];
        String sourceCellAddress = args[3];

        List<String> fileNames;
        try {
            fileNames = FileNamesReader.read(fileNamesFilePath, sourceCellAddress);
            System.out.println("file names read from " + fileNamesFilePath);
        } catch (FileNotFoundException e) {
            System.err.println("file " + fileNamesFilePath + " does not exist");
            return;
        } catch (IOException e) {
            System.err.println("file " + fileNamesFilePath + " exists but cannot be read");
            return;
        } catch (InvalidFormatException e) {
            System.err.println("file " + fileNamesFilePath + " has invalid format");
            return;
        }

        Workbook workbook;
        try (InputStream inputStream = new FileInputStream(templateFilePath)) {
            workbook = WorkbookFactory.create(inputStream);
        } catch (FileNotFoundException e) {
            System.err.println("file " + templateFilePath + " does not exist");
            return;
        } catch (IOException e) {
            System.err.println("file " + templateFilePath + " exists but cannot be read");
            return;
        } catch (InvalidFormatException e) {
            System.err.println("file " + templateFilePath + " has invalid format");
            return;
        }

        String currentDirectory = System.getProperty("user.dir");
        FileGenerator fileGenerator = new FileGenerator(workbook, destinationCellAddress, currentDirectory);
        for (String fileName : fileNames) {
            try {
                fileGenerator.generate(fileName);
                System.out.println("saved file " + fileName);
            } catch (IOException e) {
                System.err.println("could not save file " + fileName);
            }
        }
    }

}
