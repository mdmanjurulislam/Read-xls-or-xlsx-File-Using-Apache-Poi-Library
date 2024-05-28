package com.excelfileread.service;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

@Service
public class xlsAndxlsxFileRead {

    public final int englishNameField = 3;
    public final int japaneseNameField = 2;
    public final int akaField = 4; //

    // List of sheet indices where "Also Known As" column is not present
    public final List<Integer> noAkaColumnSheets = Arrays.asList(1,8,11,18,19,20,24, 25, 27, 31, 33, 34,39);

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

            // Start from the third row
            for (int r = 2; r <= rows; r++) {
                Row row = sheet.getRow(r);

                if (row == null) {
                    continue;
                }

                Cell englishCell = row.getCell(englishNameField);
                Cell japaneseCell = row.getCell(japaneseNameField);
                Cell akaCell = null;

                if (!noAkaColumnSheets.contains(i)) {
                    akaCell = row.getCell(akaField);
                }

                // Process Japanese Notation cell
                if (japaneseCell != null && japaneseCell.getCellType() == CellType.STRING) {
                    System.out.println(japaneseCell.getAddress() + "| Japanese Notation: " + japaneseCell.getStringCellValue() + "\t");
                }

                // Process English Notation cell
                if (englishCell != null && englishCell.getCellType() == CellType.STRING) {
                    System.out.println(englishCell.getAddress() + "| English Notation: " + englishCell.getStringCellValue() + "\t");
                }

                // Process Also Known As cell or handle no AKA column sheets
                if (noAkaColumnSheets.contains(i)) {
                    System.out.println("No alias name found\t\n");
                } else if (akaCell != null && akaCell.getCellType() == CellType.STRING) {
                    System.out.println(akaCell.getAddress() + "| Also Known As: " + akaCell.getStringCellValue() + "\t\n");
                } else{
                    System.out.print("NULL\t\n");
                }

                System.out.println();
            }
            System.out.println();
        }

        workbook.close();
        inputStream.close();
    }
}
