package mainFlow;

import common.FileHandler;
import dataStructure.ValidGoal;
import msexcel.ExcelCell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import validate.MyComparision;

import java.io.File;
import java.io.FilenameFilter;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import static test.ExcelForRu.proj_path;

public class VBS {
    public final static String vbsFileName = "_vbsFile";
    public final static String vbsExcelName = "_result";

    public static void execVBSFile(String vbsfileNameWithPath) throws IOException {
        Runtime.getRuntime().exec("cscript " + vbsfileNameWithPath);
    }

    public static void execAllVBSFiles(String dir) throws IOException {
        File f = new File(dir);
        File[] matchingFiles = f.listFiles(new FilenameFilter() {
            public boolean accept(File dir, String name) {
                return name.endsWith("vbs");
            }
        });
        for (File vbsfiles : matchingFiles) {
            System.out.println("execute " + vbsfiles.getPath());
            Runtime.getRuntime().exec("cscript " + vbsfiles.getPath());
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
    public static HashMap<String, ExcelCell> storeTargetInFile(Sheet sheet, HashMap<String, ValidGoal> TobeProcessed) {

        int lastRowIdx = sheet.getLastRowNum();

        //output r1c1 , temp target cell
        HashMap<String, ExcelCell> allTgt = new HashMap<String, ExcelCell>();
        for (Map.Entry<String, ValidGoal> e : TobeProcessed.entrySet()) {
            Row row = sheet.createRow(++lastRowIdx);
            int tgtCellIdx = -1;
            ValidGoal data = e.getValue();
            if (data != null) {
                MyComparision myTgts = data.getMyComparision();
                if (myTgts.OOSNum != null && myTgts.OOSNum.length > 1) {
                    if (myTgts.getTargetSmall() && myTgts.getTragetBig()) {
                        for (double num : myTgts.OOSNum) {
                            row.createCell(++tgtCellIdx).setCellValue(num);
                            allTgt.put(e.getKey() + "_" + tgtCellIdx, new ExcelCell(row.getCell(tgtCellIdx)));
                        }
                    } else {
                        row.createCell(++tgtCellIdx).setCellValue(myTgts.OOSNum[0]);
                        allTgt.put(e.getKey() + "_" + tgtCellIdx, new ExcelCell(row.getCell(tgtCellIdx)));
                    }
                }
            }

        }
//        excel.save();
//        excel.saveToFile(FileName);
        return allTgt;
    }

    public static void completeVBSContent(StringBuilder vbsFileContent, String FileName, String newFilePath) {
        vbsFileContent.insert(0, "Dim objExcel" + System.getProperty("line.separator") +
                "Dim xlSheet" + System.getProperty("line.separator") +
                "Set objExcel = CreateObject(\"Excel.Application\")" + System.getProperty("line.separator") +
                "Set objWorkbook = objExcel.Workbooks.Open(\"" + proj_path + FileName + "\")" + System.getProperty("line.separator")
        );

        vbsFileContent.append("objExcel.ActiveWorkbook.SaveAs \"" + newFilePath + "\"\n" +
                "objExcel.ActiveWorkbook.Close\n" +
                "objExcel.Application.Quit");
    }

    public static void produceNewExcels(String FileName, Sheet sheet, HashMap<String, ValidGoal> TobeProcessed) {
        String sheetName = sheet.getSheetName();

        HashMap<String, ExcelCell> allTarget = storeTargetInFile(sheet, TobeProcessed);

        for (Map.Entry<String, ExcelCell> target : allTarget.entrySet()) {
            StringBuilder vbsFileContent = new StringBuilder();
            String outputR1C1_index[] = target.getKey().split("_");
            if (outputR1C1_index.length > 1) {
                String index = outputR1C1_index[1];
                String output = outputR1C1_index[0];
                if (TobeProcessed.containsKey(output)) {
                    String newPath = proj_path + FileHandler.getFileNameWoExt(FileName) +
                            "/" + sheetName +
                            "/";
                    String newExcelPath = newPath + output + index + FileName.substring(FileName.indexOf("."), FileName.length());
                    String newVBSPath = newPath + output + index + ".vbs";
                    ValidGoal data = TobeProcessed.get(output);
                    vbsFileContent.append(saveVBSFile(data.getInput().getR1c1(), data.getOutput().getR1c1(), target.getValue().getR1c1()));
                    completeVBSContent(vbsFileContent, FileName, newExcelPath);
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
