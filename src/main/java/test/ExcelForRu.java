package test;


import common.FileHandler;
import dataStructure.ValidGoal;
import mainFlow.VBS;
import mainFlow.extract;
import msexcel.Excel;
import msexcel.ExcelCell;
import msword.CustomTableText;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import javax.swing.filechooser.FileSystemView;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;


import static mainFlow.ProduceWordFile.writeToWord_general;
import static mainFlow.ProduceWordFile.writeToWord_testCase;
import static mainFlow.VBS.execAllVBSFiles;
import static msexcel.Excel.*;
import static msexcel.customExcelStyle.setMyConditionalFormatting;

import static msword.CustomTableText.getFormulaCellAddress;
import static msword.CustomTableText.getNonFormulaCellAddress;
import static msword.CustomTableText.getOutputCellAddress;
import static validate.excelFormulaValidator.getValidatedValues;


/*
執行前要: 1.檔案要存在桌面 2.檔案名稱不能有空格&中文!!!! (CMD無法執行，VBS會無法產新excel)
[Extract]
    1.找規格(=目標值)
    2.找輸入格
    3.找輸出格
[Word]
    3.1. 產生word檔案(general)
 [excel formatting]
  4.0 更改excel 輸出值
[VBS]
    4.將目標值存入資料表
    5.產出vbs檔案
    6.執行vbs (會執行goal seek產出理想輸入格--->存成新資料表)
        -目前: 一個sheet(多個spec)-->一個vbs
        -方案: 一個目標-->一個vbs (特定命名方式= output的r1c1+目標的index)
[Word]
    7.從新資料表讀入特定VBS產出的資料表(特定命名方式) **只留下該sheet
    8.產生word檔案(test case)
 */


public class ExcelForRu {
    public static final String proj_path = FileSystemView.getFileSystemView().getHomeDirectory().getPath() + File.separator ;

    public static void main(String[] args) throws IOException, InterruptedException
//            ,RangeException
    {
//        String output = "RT30354-LAB_Spreadsheet.xlsx";
        String output = "C-RT30358LABSpreadsheet.xlsx";
        Excel excel = Excel.loadExcel(proj_path + output);

if (excel==null)   System.out.println("NULL!");
        List<Excel> excelList = dividExcelBySheet(excel,proj_path);


        XWPFDocument doc_general = new XWPFDocument();
        XWPFDocument doc_testCase = new XWPFDocument();

//        int sheetSize = Excel.getNumberOfSheets();
 int count =0;
        for (Excel outputExcel: excelList) {
            outputExcel.assignSheet(0);

             //outputExcel.assignSheet(a);
            String sheetName = outputExcel.getSheet().getSheetName();

            String sheetName_forNewFiles = outputExcel.getSheet().getSheetName().replaceAll(" ", "");
            String vbs_newData_path = proj_path + FileHandler.getFileNameWoExt(output) + "/" + sheetName_forNewFiles + "/"+"VBS"+ "/";
            String excel_newData_path = proj_path + FileHandler.getFileNameWoExt(output) + "/" + sheetName_forNewFiles + "/";

            if (!sheetName.toLowerCase().equals("history of versions")) {
                HashMap<String, ValidGoal> TobeProcessed =   extract.extractData(outputExcel);

                int validGoalsNumberInSheet = TobeProcessed.size();

//set conditional formatting for all outpull cells in this sheet
                outputExcel = setMyConditionalFormatting(TobeProcessed, outputExcel);
                outputExcel.save();
//                CustomTableText customTableText =new CustomTableText();
                CustomTableText customTableText =new CustomTableText(outputExcel,TobeProcessed);
                writeToWord_general(sheetName, doc_general, TobeProcessed,customTableText);
//                writeToWord_general(sheetName, doc_general, TobeProcessed);

//                store target goal in existing excelForRead, because vbs function--'goal seek' requires an object
                HashMap<String, ExcelCell> allTarget = VBS.storeTargetInFile(outputExcel, TobeProcessed);
                outputExcel.save();
                VBS.produceVBSFiles(String.valueOf(count), output,outputExcel.getSheet(), excel_newData_path,vbs_newData_path, allTarget, TobeProcessed);
                execAllVBSFiles(vbs_newData_path);
//                //get new Excel everytime vbs file produce new file
                HashMap<String, ValidGoal> newData = getValidatedValues(outputExcel, sheetName, TobeProcessed, excel_newData_path);



                int testCaseIdx = count + 2;


//                //要把hashmap裡面的Key改成OutputR1C1 + index -->因為有值會有上下標，需要有兩個新excel檔案!!
                writeToWord_testCase(doc_testCase, TobeProcessed, newData, testCaseIdx,customTableText);
//
                FileHandler.save(doc_general, proj_path + sheetName + "_result.docx");
                FileHandler.save(doc_testCase, proj_path + sheetName + "_test case.docx");
                count++;
            }
        }


        System.out.println("finish");


    }
}
