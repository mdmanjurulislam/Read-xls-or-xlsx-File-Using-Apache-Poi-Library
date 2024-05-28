package com.excelfileread.service;

import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.IOException;

@Service
public class xlsxFileRead {


    public String excelFileRead(String filepath) throws IOException {

        FileInputStream inputStream = new FileInputStream(filepath);

        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        XSSFSheet sheet = workbook.getSheetAt(0);

        int rows = sheet.getLastRowNum();
        int cols = sheet.getRow(1).getLastCellNum();

        for (int r = 0;r<=rows;r++){

            XSSFRow row = sheet.getRow(r);

            for(int c = 0;c<cols;c++){
                XSSFCell cell = row.getCell(c);
                if(r == 0){
                    continue;
                }else {
                    switch (cell.getCellType()) {
                        case STRING:
                            System.out.println(cell.getStringCellValue());
                            break;
                        case NUMERIC:
                            System.out.println(cell.getNumericCellValue());
                            break;
                        case BOOLEAN:
                            System.out.println(cell.getBooleanCellValue());
                            break;
                    }
                }
            }
            System.out.println();
        }

        return "File Read successfully";

    }

}
