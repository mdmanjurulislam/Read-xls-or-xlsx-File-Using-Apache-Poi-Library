package com.excelfileread.service;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Objects;

@Service
public class ParseJapaneseFileData {

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

//  ddUNSC = Date designated by the United Nations Sanctions Committee
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

        // Iterate through all the sheets in the workbook
        for (int i = 1; i < workbook.getNumberOfSheets(); i++) {

            //Iterate all the sheet from the file
            Sheet sheet = workbook.getSheetAt(i);

            //storing sheet name
            String sheetName = sheet.getSheetName();
            System.out.println("Reading sheet: " + sheetName);


            if(Objects.equals(sheetName, "1.ミロシェビッチ前ユーゴ大統領関係者")){
                System.out.println("This is sheet : 1");

                for(int rowIndex = 2;rowIndex<=sheet.getLastRowNum();rowIndex++){
                    Row row = sheet.getRow(rowIndex);

                    if(row == null){
                        continue;
                    }

                    Cell cell = row.getCell(0);
                    noticeDate = cell != null ? cell.toString() : null;

                    cell = row.getCell(1);
                    noticeNumber = cell != null ? cell.toString() : null;

                    cell = row.getCell(2);
                    japaneseNotation = cell != null ? cell.toString() : null;

                    cell = row.getCell(3);
                    englishNotation = cell != null ? cell.toString() : null;

                    cell = row.getCell(4);
                    DobAndPob = cell != null ? cell.toString() : null;

                    cell = row.getCell(5);
                    otherInformation = cell != null ? cell.toString() : null;


                    System.out.println("Notice Date         : " + noticeDate);
                    System.out.println("Notice Number       : " + noticeNumber);
                    System.out.println("Japanese Notation   : " + japaneseNotation);
                    System.out.println("English Notation    : " + englishNotation);
                    System.out.println("DOB and POB         : " + DobAndPob);
                    System.out.println("Other Information   : " + otherInformation);

                    System.out.println("==================================================");
                }

            }else if (Objects.equals(sheetName, "2.タリバーン関係者等")) {
                System.out.println("This is sheet : 2");

                for (int rowIndex = 2; rowIndex <= sheet.getLastRowNum(); rowIndex++){
                    Row row = sheet.getRow(rowIndex);

                    if(row == null){
                        continue;
                    }

                    Cell cell = row.getCell(0);
                    noticeDate = cell != null ? cell.toString() : null;

                    cell = row.getCell(1);
                    noticeNumber = cell != null ? cell.toString() : null;

                    cell = row.getCell(2);
                    japaneseNotation = cell != null ? cell.toString() : null;

                    cell = row.getCell(3);
                    englishNotation = cell != null ? cell.toString() : null;

                    cell = row.getCell(4);
                    aka = cell != null ? cell.toString() : null;

                    cell = row.getCell(5);
                    oldName = cell != null ? cell.toString() : null;

                    cell = row.getCell(6);
                    title = cell != null ? cell.toString() : null;

                    cell = row.getCell(7);
                    post = cell != null ? cell.toString() : null;

                    cell = row.getCell(8);
                    DOB = cell != null ? cell.toString() : null;

                    cell = row.getCell(9);
                    POB = cell != null ? cell.toString() : null;

                    cell = row.getCell(9);
                    nationality = cell != null ? cell.toString() : null;

                    cell = row.getCell(10);
                    passportNumber = cell != null ? cell.toString() : null;

                    cell = row.getCell(11);
                    idNumber = cell != null ? cell.toString() : null;

                    cell = row.getCell(12);
                    addressOrLocation = cell != null ? cell.toString() : null;

//                  ddUNSC = Date designated by the United Nations Sanctions Committee
                    cell = row.getCell(13);
                    ddUNSC = cell != null ? cell.toString() : null;

                    cell = row.getCell(14);
                    otherInformation = cell != null ? cell.toString() : null;


//                    printing all cell value of sheet 2
                    printSheet2and3();

//                    System.out.println("Notice Date         : " + noticeDate);
//                    System.out.println("Notice Number       : " + noticeNumber);
//                    System.out.println("Japanese Notation   : " + japaneseNotation);
//                    System.out.println("English Notation    : " + englishNotation);
//                    System.out.println("Also Known As       : " + aka);
//                    System.out.println("Old Name            : " + oldName);
//                    System.out.println("Title               : " + title);
//                    System.out.println("Date of Birth       : " + DOB);
//                    System.out.println("Place of Birth      : " + POB);
//                    System.out.println("Nationality         : " + nationality);
//                    System.out.println("Passport Number     : " + passportNumber);
//                    System.out.println("ID Number           : " + idNumber);
//                    System.out.println("Address/Location    : " + addressOrLocation);
//                    System.out.println("Date Designed by United Sanction Committee : " + ddUNSC);
//                    System.out.println("Other Information   : " + otherInformation);

                    System.out.println("==================================================");
                }

            } else if (Objects.equals(sheetName, "3.テロリスト等 (1)")) {
                System.out.println("This is Sheet : 3");

                for (int rowIndex = 3; rowIndex <= sheet.getLastRowNum(); rowIndex++){
                    Row row = sheet.getRow(rowIndex);

                    if(row == null){
                        continue;
                    }

                    Cell cell = row.getCell(0);
                    noticeDate = cell != null ? cell.toString() : null;

                    cell = row.getCell(1);
                    noticeNumber = cell != null ? cell.toString() : null;

                    cell = row.getCell(2);
                    japaneseNotation = cell != null ? cell.toString() : null;

                    cell = row.getCell(3);
                    englishNotation = cell != null ? cell.toString() : null;

                    cell = row.getCell(4);
                    aka = cell != null ? cell.toString() : null;

                    cell = row.getCell(5);
                    oldName = cell != null ? cell.toString() : null;

                    cell = row.getCell(6);
                    title = cell != null ? cell.toString() : null;

                    cell = row.getCell(7);
                    post = cell != null ? cell.toString() : null;

                    cell = row.getCell(8);
                    DOB = cell != null ? cell.toString() : null;

                    cell = row.getCell(9);
                    POB = cell != null ? cell.toString() : null;

                    cell = row.getCell(10);
                    nationality = cell != null ? cell.toString() : null;

                    cell = row.getCell(11);
                    passportNumber = cell != null ? cell.toString() : null;

                    cell = row.getCell(12);
                    idNumber = cell != null ? cell.toString() : null;

                    cell = row.getCell(13);
                    addressOrLocation = cell != null ? cell.toString() : null;

//                  ddUNSC = Date designated by the United Nations Sanctions Committee
                    cell = row.getCell(14);
                    ddUNSC = cell != null ? cell.toString() : null;

                    cell = row.getCell(15);
                    otherInformation = cell != null ? cell.toString() : null;



//                    printing all cell value of sheet 3
                    printSheet2and3();

//                    System.out.println("Notice Date         : " + noticeDate);
//                    System.out.println("Notice Number       : " + noticeNumber);
//                    System.out.println("Japanese Notation   : " + japaneseNotation);
//                    System.out.println("English Notation    : " + englishNotation);
//                    System.out.println("Also Known As       : " + aka);
//                    System.out.println("Old Name            : " + oldName);
//                    System.out.println("Title               : " + title);
//                    System.out.println("Date of Birth       : " + DOB);
//                    System.out.println("Place of Birth      : " + POB);
//                    System.out.println("Nationality         : " + nationality);
//                    System.out.println("Passport Number     : " + passportNumber);
//                    System.out.println("ID Number           : " + idNumber);
//                    System.out.println("Address/Location    : " + addressOrLocation);
//                    System.out.println("Date Designed by United Sanction Committee : " + ddUNSC);
//                    System.out.println("Other Information   : " + otherInformation);

                    System.out.println("==================================================");
                }
            }

            System.out.println();
        }

        workbook.close();
        inputStream.close();
    }

    public void printSheet2and3(){
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
        System.out.println("Date Designed by United Sanction Committee : " + ddUNSC);
        System.out.println("Other Information   : " + otherInformation);
    }
}
