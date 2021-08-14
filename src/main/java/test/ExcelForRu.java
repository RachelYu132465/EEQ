package test;

import msexcel.Excel;
import org.apache.poi.ss.usermodel.Cell;

import javax.swing.filechooser.FileSystemView;
import java.io.IOException;


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
        -方案: 一個目標-->一個vbs
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

        Excel test = Excel.loadExcel(ExcelForRu.proj_path + "/testR.xlsx");
        Cell c =test.assignSheet(0).assignRow(0).assignCell(2).getCurCell();
        System.out.println(test.getCellValue() + c.getCellType().name() );
//GeneralSTRING
//        String fileName = "test1.xlsx";


//        XWPFDocument doc_general = null;
//        XWPFDocument doc_testCase = null;
//
//                //for(Sheet sheet:sheets)
//        List<String> SpecificationList = null;
//        HashMap<String, ValidGoal> TobeProcessed = null;
//        Object result = extract.extractData(fileName);
//
//        SpecificationList = (List<String>) ((Object[]) result)[0];
//        TobeProcessed = (HashMap<String, ValidGoal>) ((Object[]) result)[1];
//
////        for(String specification : SpecificationList){
////            System.out.println(specification);
////        }
//
//        for (Map.Entry<String, ValidGoal> e : TobeProcessed.entrySet()) {
////            HashSet<ExcelCell> allin = e.getValue().getAllInputs();
////
////            for(ExcelCell c:allin){
////                System.out.print(Excel.getR1C1Idx(c.getCell())+",");
////            }
////                    e.getValue().getInput();
//            System.out.println("____" + e.getValue());
//        }
//
////        VBS.execVBSFile(VBS.produceVBSFile(fileName,TobeProcessed));
////        Thread.sleep(3000);
////
////        String produced_vbsFileName =fileName.split("\\.")[0] + VBS.vbsExcelName + "." + fileName.split("\\.")[1];
////        Excel producedExcel =Excel.loadExcel(proj_path + produced_vbsFileName);
////        producedExcel.assignSheet(0);
////        HashMap<String, ValidGoal> Validated = excelFormulaValidator.getValidatedValues(
////                TobeProcessed,producedExcel);
//
//        doc_general = ProduceWordFile.writeToWord(TobeProcessed,Validated);
//        System.out.println("finish");


    }
}
