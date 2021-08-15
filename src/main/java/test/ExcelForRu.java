package test;


import common.FileHandler;
import dataStructure.ValidGoal;
import mainFlow.ProduceWordFile;
import mainFlow.VBS;
import mainFlow.extract;
import msexcel.Excel;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import javax.swing.filechooser.FileSystemView;
import java.io.IOException;
import java.util.HashMap;

import static mainFlow.ProduceWordFile.writeToWord_general;
import static mainFlow.ProduceWordFile.writeToWord_testCase;
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
    7.從新資料表讀入特定VBS產出的資料表(特定命名方式)
    8.產生word檔案(test case)
 */

/*
重複?
more or less尚未寫
description產生器
最後的word調整
*/

public class ExcelForRu {
    public static final String proj_path = FileSystemView.getFileSystemView().getHomeDirectory().getPath() + "/";

    public static void main(String[] args) throws IOException, InterruptedException {


//        Excel test = Excel.loadExcel(ExcelForRu.proj_path + "/testR.xlsx");
//        Cell c =test.assignSheet(0).assignRow(0).assignCell(2).getCurCell();

        String fileName = "test1.xlsx";
//        String fileName = "R000012383-LAB Spreadsheet數字.xlsx";
        Excel excel = Excel.loadExcel(proj_path + fileName);
        XWPFDocument doc_general = new XWPFDocument();
        XWPFDocument doc_testCase = new XWPFDocument();

        int sheetSize = excel.getNumberOfSheets();

        for (int a = 0; a < sheetSize; a++) {
            excel.assignSheet(a);
            if(!excel.getSheet().getSheetName().toLowerCase().equals("history of versions")){
                HashMap<String, ValidGoal> TobeProcessed = extract.extractData(excel);
//                writeToWord_general(excel.getSheet().getSheetName(),doc_general,TobeProcessed);
//                VBS.produceNewExcels(fileName, excel.getSheet(),TobeProcessed);
//                excel.save();
                //get new Excel everytime vbs file produce new file
                HashMap<String, ValidGoal> newData= getValidatedValues(TobeProcessed,excel);
                writeToWord_testCase(doc_testCase,TobeProcessed,newData);
            }
        }

        FileHandler.save(doc_general,proj_path+"result.docx");
        FileHandler.save(doc_testCase,proj_path+"test case.docx");
//        for (Map.Entry<String, ValidGoal> e : TobeProcessed.entrySet()) {
//            System.out.println("____" +  e.getValue());
//        }


//        VBS.execVBSFile(VBS.produceVBSFile(fileName,TobeProcessed));
//        Thread.sleep(3000);


        System.out.println("finish");


    }
}
