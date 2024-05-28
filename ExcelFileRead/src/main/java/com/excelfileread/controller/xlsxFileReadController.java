package com.excelfileread.controller;

import com.excelfileread.service.*;
import lombok.AllArgsConstructor;
import lombok.NoArgsConstructor;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;

@AllArgsConstructor
@NoArgsConstructor
@RestController
public class xlsxFileReadController {

    @Autowired
    private xlsxFileRead xlsxFileRead;

    @Autowired
    private xlsAndxlsxFileRead xlsAndxlsxFile;

    @Autowired
    private ReadColFromxls readColFromxls;

    @Autowired
    private ParseJapaneseFileData parseJapaneseFileData;

    @Autowired
    private ParseJapanFile parseJapanFile;




    //String filePath = "C:\\Users\\manjurul.sohag\\Downloads\\translated.xls";


    String filePath = "C:\\Users\\manjurul.sohag\\Downloads\\source.xls";

    @GetMapping("/hit")
    public String excelFileReader() throws IOException {
        xlsxFileRead.excelFileRead(filePath);
        return "Done";
    }

    @GetMapping("/hit2")
    public String xlsFileRead() throws IOException {
        xlsAndxlsxFile.excelFileRead(filePath);
        return "Done";
    }

    @GetMapping("/hit3")
    public String readCol() throws IOException {
        readColFromxls.excelFileRead(filePath);
        return "Done";
    }

    @GetMapping("/hit4")
    public String allSheet() throws IOException {
        parseJapaneseFileData.excelFileRead(filePath);
        return "Done";
    }

    @GetMapping("/hit5")
    public String parseJpnFile() throws IOException {
        parseJapanFile.excelFileRead(filePath);
        return "Done";
    }


}
