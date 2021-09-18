package validate;

import common.FileHandler;
import msexcel.Excel;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.ss.usermodel.Cell;

import javax.swing.filechooser.FileSystemView;
import java.io.File;
import java.io.IOException;
import java.util.List;

import static common.FileHandler.getFilesByKeywords;
import static msexcel.Excel.getMergedRegions;

public class customExcelMicro {
    public static final String proj_path = FileSystemView.getFileSystemView().getHomeDirectory().getPath() + "/";

    public static void main(String[] args) {
        String fileName = "R000012383-LAB Spreadsheet數字.xlsx";
        Excel excel = Excel.loadExcel(proj_path + fileName);
        excel.assignSheet(0).assignRow(0).assignCell(3);
//        System.out.println(getMergedRegions(excel.getSheet(),3));
        getMergedRegions (excel.getCurCell());
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
    }}
