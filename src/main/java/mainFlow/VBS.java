package mainFlow;

import common.FileHandler;
import dataStructure.ValidGoal;
import msexcel.Excel;
import msexcel.ExcelCell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import validate.MyRange;

import validate.customStringFormatter;

import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static test.ExcelForRu.proj_path;

public class VBS {
    public final static String vbsFileName = "_vbsFile";
    public final static String vbsExcelName = "_result";

    public static void main(String... arg) throws IOException {
//        execAllVBSFiles("C:\\Users\\tina.yu\\Desktop\\RT30397\\Item 2.3");
        Runtime.getRuntime().exec("cscript C:\\Users\\tina.yu\\Desktop\\RT30397\\Item2.3\\C21_0.vbs");

        System.out.println("finish");
    }

    public static void execVBSFile(String vbsfileNameWithPath) throws IOException {
        Runtime.getRuntime().exec("cscript " + vbsfileNameWithPath);
    }

    public static void execAllVBSFiles(String dir) {
        List<File> matchingFiles = FileHandler.getFilesByKeywords(dir, "vbs");

        for (File vbsfiles : matchingFiles) {
            System.out.println("execute " + vbsfiles.getPath());
            try {
                Runtime.getRuntime().exec("cscript " + vbsfiles.getPath());
                Thread.sleep(10000);
            } catch (IOException e) {
                e.printStackTrace();
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        }

    }


    public static String saveVBSFile(String inputIdx, String outputIdx, String tgtIdx) {
        return "Dim changingCell" + System.getProperty("line.separator") +
                "Dim goal" + System.getProperty("line.separator") +
                "Set changingCell" + " = objExcel.Range(\"" + inputIdx + "\")" + System.getProperty("line.separator") +
                "Set goal" + " = objExcel.Range(\"" + tgtIdx + "\")" + System.getProperty("line.separator") +
                "Call objExcel.Range(\"" + outputIdx + "\").GoalSeek(goal" + ", changingCell" + ")" + System.getProperty("line.separator");

    }

    //before vbs file is created, store target values in the original files
    public static HashMap<String, ExcelCell> storeTargetInFile(Excel excel, HashMap<String, ValidGoal> TobeProcessed)
//            throws RangeException
    {
        Sheet sheet = excel.getSheet();
        int lastRowIdx = sheet.getLastRowNum();

        //output r1c1 , temp target cell
        HashMap<String, ExcelCell> allTgt = new HashMap<String, ExcelCell>();
        for (Map.Entry<String, ValidGoal> e : TobeProcessed.entrySet()) {
            System.out.println("last row:" + lastRowIdx);
            ValidGoal data = e.getValue();
            if (data != null) {
                MyRange myTgts = data.getMyRange();
                List<Double> OOSNum = customStringFormatter.setOOS(myTgts);
                /*
                要改!!
                 */
                if (OOSNum != null && OOSNum.size() > 0) {
                    Row row = sheet.createRow(++lastRowIdx);
                    System.out.println("create row:" + data.getOutput().getR1c1());
                    int tgtCellIdx = -1;
                    if (myTgts.hasMax() && myTgts.hasMin()) {
                        for (double num : OOSNum) {
                            row.createCell(++tgtCellIdx).setCellValue(num);
                            allTgt.put(e.getKey() + "_" + tgtCellIdx, new ExcelCell(excel,row.getCell(tgtCellIdx)));
                        }
                    } else {
                        row.createCell(++tgtCellIdx).setCellValue(OOSNum.get(0));
                        allTgt.put(e.getKey() + "_" + tgtCellIdx, new ExcelCell(excel,row.getCell(tgtCellIdx)));
                    }
                }
            }

        }
//        excel.save();
//        excel.saveToFile(FileName);
        return allTgt;
    }

    public static void completeVBSContent(StringBuilder vbsFileContent, String FileName, String sheetName, String newFilePath) {
        vbsFileContent.insert(0, "Dim objExcel" + System.getProperty("line.separator") +
                "Dim xlSheet" + System.getProperty("line.separator") +
                "Set objExcel = CreateObject(\"Excel.Application\")" + System.getProperty("line.separator") +
                "Set objWorkbook = objExcel.Workbooks.Open(\"" + proj_path + FileName + "\")" + System.getProperty("line.separator") +
                "objExcel.Sheets(\"" + sheetName + "\").Select" + System.getProperty("line.separator"));


        vbsFileContent.append("objExcel.ActiveWorkbook.SaveAs \"" + newFilePath + "\"\n" +
                "objExcel.ActiveWorkbook.Close\n" +
                "objExcel.Application.Quit");
    }

    public static void produceVBSFiles(String FileName, Sheet sheet, String newexcelPath, String newvbsPath, HashMap<String, ExcelCell> allTarget, HashMap<String, ValidGoal> TobeProcessed)
//            throws RangeException
    {
        String sheetName = sheet.getSheetName();

        for (Map.Entry<String, ExcelCell> target : allTarget.entrySet()) {
            StringBuilder vbsFileContent = new StringBuilder();
            String outputR1C1_index[] = target.getKey().split("_");
            if (outputR1C1_index.length > 1) {
                String index = outputR1C1_index[1];
                String output = outputR1C1_index[0];
                if (TobeProcessed.containsKey(output)) {
                    String newExcelPath = newexcelPath + output + "_" + index + FileName.substring(FileName.indexOf("."), FileName.length());
                    String newVBSPath = newvbsPath + output + "_" + index + ".vbs";
                    ValidGoal data = TobeProcessed.get(output);
                    vbsFileContent.append(saveVBSFile(data.getInput().getR1c1(), data.getOutput().getR1c1(), target.getValue().getR1c1()));
                    completeVBSContent(vbsFileContent, FileName, sheetName, newExcelPath);
                    FileHandler.create_save(newVBSPath, vbsFileContent.toString());
                }
            }
        }

    }

//    public static String produceVBSFile(String FileName, HashMap<String, ValidGoal> TobeProcessed) {
//
//        String fileName_noExt = FileName.split("\\.")[0];
//
//        Excel excel = Excel.loadExcel(proj_path + FileName);
//        excel.assignSheet(0);
//        int lastRowIdx = excel.getLastRowNum();
//        excel.assignRow(lastRowIdx + 1);
//        StringBuilder vbsFileContent = new StringBuilder();
//        int tgtCellIdx = -1;
//        //output r1c1 , temp target cell
//        HashMap<String, ExcelCell> allTgt = new HashMap<String, ExcelCell>();
//        for (Map.Entry<String, ValidGoal> e : TobeProcessed.entrySet()) {
//            ValidGoal data = e.getValue();
//            if (data != null) {
//                //每個output都要另存目標值在最下面一列
//                excel.assignCell(++tgtCellIdx);
//                excel.setCellValue(data.getMyComparision().getOOSNum()[0]);
////                allTgt.put(getR1C1Idx(data.getOutput().getCell()),
////                        new ExcelCell(
////                                getR1C1Idx(excel.getCurCell()),
////                                String.valueOf(data.getTargetOutput()),
////                                excel.getCurCell()));
//                vbsFileContent.append(saveVBSFile(tgtCellIdx, data.getInput().getR1c1(), data.getOutput().getR1c1(), getR1C1Idx(excel.getCurCell())));
////                System.out.println(e.getValue());
//            }
//        }
//        vbsFileContent.insert(0, "Dim objExcel" + System.getProperty("line.separator") +
//                "Dim xlSheet" + System.getProperty("line.separator") +
//                "Set objExcel = CreateObject(\"Excel.Application\")" + System.getProperty("line.separator") +
//                "Set objWorkbook = objExcel.Workbooks.Open(\"" + proj_path + FileName + "\")" + System.getProperty("line.separator")
//        );
//
//        vbsFileContent.append("objExcel.ActiveWorkbook.SaveAs \"" + proj_path + fileName_noExt + vbsExcelName + excel.getExcelType().getValue() + "\"\n" +
//                "objExcel.ActiveWorkbook.Close\n" +
//                "objExcel.Application.Quit");
//
//        try {
//            File VBSfile = new File(proj_path + fileName_noExt + vbsFileName + ".vbs");
//
//            if (VBSfile.createNewFile()) {
//                System.out.println("File created: " + VBSfile.getName());
//            } else {
//                System.out.println("File already exists.");
//            }
//            FileWriter myWriter = new FileWriter(VBSfile);
//            myWriter.write(vbsFileContent.toString());
//            myWriter.close();
//            System.out.println("Successfully wrote to the file.");
//        } catch (
//                IOException e) {
//            System.out.println("An error occurred.");
//            e.printStackTrace();
//        }
//
//        excel.saveToFile(FileName);
//        return proj_path + fileName_noExt + vbsFileName + ".vbs";
//    }

}
