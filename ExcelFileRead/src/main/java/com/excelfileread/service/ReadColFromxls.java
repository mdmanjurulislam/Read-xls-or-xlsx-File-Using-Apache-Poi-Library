package com.excelfileread.service;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.IOException;

@Service
public class ReadColFromxls {

    public final int englishNotation = 3;

    public void excelFileRead(String filepath) throws IOException {
        FileInputStream inputStream = new FileInputStream(filepath);
        Workbook workbook;

        if (filepath.endsWith(".xlsx")) {
            workbook = new XSSFWorkbook(inputStream);
        } else if (filepath.endsWith(".xls")) {
            workbook = new HSSFWorkbook(inputStream);
        } else {
            throw new IllegalArgumentException("The specified file is not an Excel file");
        }

        // Iterate through all the sheets in the workbook
        for (int i = 1; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            System.out.println("Reading sheet: " + sheet.getSheetName());

            int rows = sheet.getLastRowNum();

            for (int r = 1; r <= rows; r++) {
                Row row = sheet.getRow(r);

                if (row == null) {
                    continue; // Skip empty rows
                }

                Cell cell = row.getCell(3); // Column D has index 3

                if (cell == null) {
                    System.out.print("NULL\t");
                } else {
                    switch (cell.getCellType()) {
                        case STRING:
                            if (cell.equals((cell = row.getCell(englishNotation)))) {
                                System.out.println(cell.getAddress() + "English Notation : " + cell.getStringCellValue() + "\t");
                            }
//                            System.out.println(cell.getAddress() + "English Notation : " + cell.getStringCellValue() + "\t");
                            break;
                        case NUMERIC:
                            System.out.println(cell.getNumericCellValue() + "\t");
                            break;
                        case BOOLEAN:
                            System.out.println(cell.getBooleanCellValue() + "\t");
                            break;
                        case FORMULA:
                            System.out.println(cell.getCellFormula() + "\t");
                            break;
                        default:
                            System.out.print("UNKNOWN\t");
                            break;
                    }
                }
                System.out.println();
            }
            System.out.println();
        }

        workbook.close();
        inputStream.close();
    }
}
