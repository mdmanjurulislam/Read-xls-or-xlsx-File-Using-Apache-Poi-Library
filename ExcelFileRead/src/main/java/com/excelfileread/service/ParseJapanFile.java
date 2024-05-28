package com.excelfileread.service;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Objects;

@Service
public class ParseJapanFile {

    public String noticeDate;
    public String noticeNumber;
    public String japaneseNotation;
    public String englishNotation;
    public String DobAndPob;
    public String otherInformation;
    public String aka;
    public String oldName;
    public String title;
    public String post;
    public String POB;
    public String DOB;
    public String nationality;
    public String passportNumber;
    public String idNumber;
    public String addressOrLocation;
    public String ddUNSC;

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

        for (int i = 1; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            String sheetName = sheet.getSheetName();
            System.out.println("Reading sheet: " + sheetName);

            if (Objects.equals(sheetName, "1.ミロシェビッチ前ユーゴ大統領関係者")) {
                processSheet1(sheet);
            } else if (Objects.equals(sheetName, "2.タリバーン関係者等") || Objects.equals(sheetName, "3.テロリスト等 (1)")) {
                processSheet2and3(sheet);
            }
        }

        workbook.close();
        inputStream.close();
    }

    private void processSheet1(Sheet sheet) {
        System.out.println("This is sheet : 1");
        for (int rowIndex = 2; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) continue;

            noticeDate = getCellValue(row, 0);
            noticeNumber = getCellValue(row, 1);
            japaneseNotation = getCellValue(row, 2);
            englishNotation = getCellValue(row, 3);
            DobAndPob = getCellValue(row, 4);
            otherInformation = getCellValue(row, 5);

            printSheet1Details();
        }
    }

    private void processSheet2and3(Sheet sheet) {
        System.out.println("This is sheet : " + sheet.getSheetName());
        for (int rowIndex = 2; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) continue;

            String[] cellValues = new String[16];
            for (int cellIndex = 0; cellIndex < cellValues.length; cellIndex++) {
                cellValues[cellIndex] = getCellValue(row, cellIndex);
            }

            noticeDate = cellValues[0];
            noticeNumber = cellValues[1];
            japaneseNotation = cellValues[2];
            englishNotation = cellValues[3];
            aka = cellValues[4];
            oldName = cellValues[5];
            title = cellValues[6];
            post = cellValues[7];
            DOB = cellValues[8];
            POB = cellValues[9];
            nationality = cellValues[10];
            passportNumber = cellValues[11];
            idNumber = cellValues[12];
            addressOrLocation = cellValues[13];
            ddUNSC = cellValues[14];
            otherInformation = cellValues[15];

            printSheet2and3Details();
        }
    }

    private String getCellValue(Row row, int cellIndex) {
        Cell cell = row.getCell(cellIndex);
        return cell != null ? cell.toString() : null;
    }

    private void printSheet1Details() {
        System.out.println("Notice Date         : " + noticeDate);
        System.out.println("Notice Number       : " + noticeNumber);
        System.out.println("Japanese Notation   : " + japaneseNotation);
        System.out.println("English Notation    : " + englishNotation);
        System.out.println("DOB and POB         : " + DobAndPob);
        System.out.println("Other Information   : " + otherInformation);
        System.out.println("==================================================");
    }

    private void printSheet2and3Details() {
        System.out.println("Notice Date         : " + noticeDate);
        System.out.println("Notice Number       : " + noticeNumber);
        System.out.println("Japanese Notation   : " + japaneseNotation);
        System.out.println("English Notation    : " + englishNotation);
        System.out.println("Also Known As       : " + aka);
        System.out.println("Old Name            : " + oldName);
        System.out.println("Title               : " + title);
        System.out.println("Date of Birth       : " + DOB);
        System.out.println("Place of Birth      : " + POB);
        System.out.println("Nationality         : " + nationality);
        System.out.println("Passport Number     : " + passportNumber);
        System.out.println("ID Number           : " + idNumber);
        System.out.println("Address/Location    : " + addressOrLocation);
        System.out.println("Date Designed by United Nations Sanctions Committee : " + ddUNSC);
        System.out.println("Other Information   : " + otherInformation);
        System.out.println("==================================================");
    }
}

