package test;


import common.FileHandler;
import dataStructure.ValidGoal;
import mainFlow.VBS;
import mainFlow.extract;
import msexcel.Excel;
import msexcel.ExcelCell;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import javax.swing.filechooser.FileSystemView;
import java.io.IOException;
import java.util.HashMap;

import static mainFlow.ProduceWordFile.writeToWord_general;
import static mainFlow.ProduceWordFile.writeToWord_testCase;
import static mainFlow.VBS.execAllVBSFiles;
import static msexcel.customExcelStyle.setMyConditionalFormatting;
import static test.test.sortCell;
import static validate.excelFormulaValidator.getValidatedValues;


/*
執行前要: 1.檔案要存在桌面 2.檔案名稱不能有空格 (CMD無法執行，VBS會無法產新excel)
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
    public static final String proj_path = FileSystemView.getFileSystemView().getHomeDirectory().getPath() + "/";

    public static void main(String[] args) throws IOException, InterruptedException
//            ,RangeException
    {
//        String fileName =  "格式_RT30397_LABSpreadsheet數字(1).xlsx";
//        String fileName = "C- RT30358-LAB Spreadsheet-數字版.xlsx";
//        String fileName = "final_RT30358_LAB_Spreadsheet複製.xlsx";
        //String fileName =  "final_RT30358_LAB_Spreadsheet_2.xlsx";
        String output =  "originSpec_RT30358_LAB_Spreadsheet.xlsx";
//        String fileName = "R000012383-LAB Spreadsheet數字.xlsx";
        Excel outputExcel = Excel.loadExcel(proj_path + output);
        //Excel excel = Excel.loadExcel(proj_path + fileName);
        System.out.println(outputExcel.assignSheet(0).getSheetName(0));

        XWPFDocument doc_general = new XWPFDocument();
        XWPFDocument doc_testCase = new XWPFDocument();

        int sheetSize = outputExcel.getNumberOfSheets();

        for (int a = 0; a < sheetSize; a++) {
            outputExcel.assignSheet(a);

             //outputExcel.assignSheet(a);
            String sheetName = outputExcel.getSheet().getSheetName();

            String sheetName_forNewFiles = outputExcel.getSheet().getSheetName().replaceAll(" ", "");
            String vbs_newData_path = proj_path + FileHandler.getFileNameWoExt(output) + "/" + sheetName_forNewFiles + "/"+"VBS"+ "/";
            String excel_newData_path = proj_path + FileHandler.getFileNameWoExt(output) + "/" + sheetName_forNewFiles + "/";

            if (!sheetName.toLowerCase().equals("history of versions")) {
                HashMap<String, ValidGoal> TobeProcessed =  sortCell(extract.extractData(outputExcel));

                int validGoalsNumberInSheet = TobeProcessed.size();

//set conditional formatting for all outpull cells in this sheet
                outputExcel = setMyConditionalFormatting(TobeProcessed, outputExcel);
                outputExcel.save();

//                writeToWord_general(sheetName, doc_general, TobeProcessed);


                //store target goal in existing excelForRead, because vbs function--'goal seek' requires an object
                HashMap<String, ExcelCell> allTarget = VBS.storeTargetInFile(outputExcel, TobeProcessed);
                outputExcel.save();
                VBS.produceVBSFiles(output, outputExcel.getSheet(), vbs_newData_path, allTarget, TobeProcessed);
                execAllVBSFiles(excel_newData_path);
//                //get new Excel everytime vbs file produce new file
                HashMap<String, ValidGoal> newData = getValidatedValues(outputExcel, sheetName, TobeProcessed, excel_newData_path);

                int testCaseIdx = a + 2;
                //要把hashmap裡面的Key改成OutputR1C1 + index -->因為有值會有上下標，需要有兩個新excel檔案!!
                writeToWord_testCase(doc_testCase, TobeProcessed, newData, testCaseIdx);

                FileHandler.save(doc_general, proj_path + sheetName + "_result.docx");
                FileHandler.save(doc_testCase, proj_path + sheetName + "_test case.docx");
            }
        }


        System.out.println("finish");


    }
}
