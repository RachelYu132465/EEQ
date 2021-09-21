package customExcel;

import common.FileHandler;
import msexcel.Excel;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;


import javax.swing.filechooser.FileSystemView;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.List;

import static common.FileHandler.getFilesByKeywords;
import static common.FileHandler.moveFileToSubFolder;
import static msexcel.Excel.*;

public class customExcelMicro {
    public static final String proj_path = FileSystemView.getFileSystemView().getHomeDirectory().getPath() + "\\09_13_21_CODE";

    public static void main(String[] args) throws IOException,InterruptedException {

//        //將有micro檔移入subfolder
//        List<File> f = getFilesByKeywords(proj_path, "");
//        for (File ff : f) {
//            String extension = FilenameUtils.getExtension(ff.getAbsolutePath());
//            if (extension.equalsIgnoreCase("xlsx") ||
//                    extension.equalsIgnoreCase("xls")) {
//                Excel excel = Excel.loadExcel(ff);
//                List<Sheet> sheet = getSheetsBykeywordsIgnoreCase(excel, "micro");
//if(sheet.size()>0){
//
//
//  excel.getWorkbook().close();
//  moveFileToSubFolder(proj_path, "micro", ff);
//}

//把版型複製入檔案


        List<File> files = getFilesByKeywords(proj_path + "\\micro", "");

//        //複製result CELL的格式
//        for (File file : files) {
//            Excel excel = Excel.loadExcel(file);
//            List<String> sheet = getSheetsNameBykeywordsIgnoreCase(excel,"micro");
//            int microShhetidx  = excel.getWorkbook().getSheetIndex(sheet.get(0));
//            excel.assignSheet(microShhetidx);
////            excel.getWorkbook().assignRow(6).assignCell(1);
//            CellStyle oldStyle = excel.getCurCell().getCellStyle();
//            excel.assignRow(12);
//            for (int i=1;i<5;i++){
//                excel.assignCell(i);
//
//
//                CellStyle style = excel.getWorkbook().createCellStyle();
//                style.setBorderBottom(BorderStyle.THIN);
//                style.cloneStyleFrom(oldStyle);
//                excel.getCurCell().setCellStyle(style);
//
//            }

//            excel.save();
//        excel.outputFile("C:\\Users\\YUY139\\Desktop\\LAB Spreadsheet");

        }

//        for (Sheet s : sheet) {
//
//
//        }

                String fileName = "R000007847-LAB Spreadsheet.xlsx";
                String fileName2 = "H000019703_Biotin & Micro.xlsx";

//            }
//        Excel excel = Excel.loadExcel(proj_path + fileName);
//        Excel excel2 = Excel.loadExcel(proj_path + fileName2);
//        excel2.copySheet()
//        excel.assignSheet(0).assignRow(1).assignCell(3);
//        excel.setCellValue("QQ");
//        excel.save();
//        excel.outputFile("C:\\Users\\YUY139\\Desktop\\R000007847-LAB Spreadsheet");
//        System.out.println(getMergedRegions(excel.getSheet(),3));
//        getMergedRegions (excel.getCurCell());


//        List<File> f = getFilesByKeywords(proj_path, "");
//        for (File ff : f) {
//            String extension = FilenameUtils.getExtension(ff.getAbsolutePath());
//
//            if (extension.equalsIgnoreCase("xlsx") ||
//                    extension.equalsIgnoreCase("xls")) {
//                Excel excel = Excel.loadExcel(ff);
//
//                int sheetSize = excel.getNumberOfSheets();
//                for (int a = 0; a < sheetSize; a++) {
//                    String sheetName = excel.getSheet().getSheetName();
//                    if (sheetName.toLowerCase().equals("micro limit test") || sheetName.toLowerCase().equals("Microbial Limit")) {
//                        excel.assignSheet(a);
//
//                        //需求一 增加version字樣
//                        String cellAddr = excel.findFirstWordInWb("version", excel);
//                        excel.setCellValue(cellAddr, "Spreadsheet Version");
//                        Cell version = excel.getCell(cellAddr);
//                        version.getAddress().getRow();
//                        version.getAddress().getColumn();
//                        int rowSize = excel.getLastRowNum();
//                        for (int j = 0; j < rowSize; j++) {
//                            excel.assignRow(j);
//                            int colSize = excel.getLastCellNum();
//                            for (int k = 0; k < colSize; k++) {
//                                excel.assignCell(k);
//                            }
//                        }
//
//                    }
//                }
//            }}
    }
