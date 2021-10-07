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

    }
    }
