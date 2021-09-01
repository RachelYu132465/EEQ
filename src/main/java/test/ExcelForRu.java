package test;


import common.FileHandler;
import dataStructure.ValidGoal;
import mainFlow.VBS;
import mainFlow.extract;
import msexcel.Excel;
import msexcel.ExcelCell;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import validate.RangeException;

import javax.swing.filechooser.FileSystemView;
import java.io.IOException;
import java.util.HashMap;

import static mainFlow.ProduceWordFile.writeToWord_general;
import static mainFlow.ProduceWordFile.writeToWord_testCase;
import static mainFlow.VBS.execAllVBSFiles;
import static validate.excelFormulaValidator.getValidatedValues;


/*
[Extract]
    1.找規格(=目標值)
    2.找輸入格
    3.找輸出格
[Word]
    3.1. 產生word檔案(general)
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

    public static void main(String[] args) throws IOException, InterruptedException, RangeException {

        String fileName = "RT30397.xlsx";
//        String fileName = "R000012383-LAB Spreadsheet數字.xlsx";
        Excel excel = Excel.loadExcel(proj_path + fileName);
        XWPFDocument doc_general = new XWPFDocument();
        XWPFDocument doc_testCase = new XWPFDocument();

        int sheetSize = excel.getNumberOfSheets();

        for (int a = 0; a < sheetSize; a++) {
            excel.assignSheet(a);
            String sheetName = excel.getSheet().getSheetName();

            String sheetName_forNewFiles = excel.getSheet().getSheetName().replaceAll(" ","");
            String vbs_newData_path = proj_path + FileHandler.getFileNameWoExt(fileName) + "/" + sheetName_forNewFiles + "/";
            if (!sheetName.toLowerCase().equals("history of versions")) {
                HashMap<String, ValidGoal> TobeProcessed = extract.extractData(excel);
                int validGoalsNumberInSheet = TobeProcessed.size();
                writeToWord_general(sheetName, doc_general, TobeProcessed);
                //store target goal in existing excel, because vbs function--'goal seek' requires an object
                HashMap<String, ExcelCell> allTarget = VBS.storeTargetInFile(excel.getSheet(), TobeProcessed);
                excel.save();
                VBS.produceVBSFiles(fileName, excel.getSheet(), vbs_newData_path, allTarget,TobeProcessed);
                execAllVBSFiles(vbs_newData_path);
                //get new Excel everytime vbs file produce new file
                HashMap<String, ValidGoal> newData = getValidatedValues(sheetName, TobeProcessed, vbs_newData_path);

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
