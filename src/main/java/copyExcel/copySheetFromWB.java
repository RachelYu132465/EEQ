package copyExcel;

import msexcel.Excel;
import org.apache.poi.hssf.usermodel.HSSFPrintSetup;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;

import javax.swing.filechooser.FileSystemView;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import static common.FileHandler.getFilesByKeywords;

public class copySheetFromWB  {

    public Excel excel;
    public static final String proj_path = FileSystemView.getFileSystemView().getHomeDirectory().getPath() + "\\09_13_21_CODE";
    public static void main(String[] args) {

        Excel template = Excel.loadExcel("C:\\Users\\YUY139\\Desktop\\template.xlsx");
        Excel excel = Excel.loadExcel("C:\\Users\\YUY139\\Desktop\\copy.xlsx");
//         copy(template,excel,0,template.getWorkbook().getSheetAt(0));



    }
    public static Excel copy (Excel template, Excel excel, int copyTo, Sheet fromsheet){


        Sheet newsheet = excel.getWorkbook().getSheetAt(0);

        excel.isDisplayGridlines(newsheet,true);
        excel.setMargin(newsheet,fromsheet.getMargin(Sheet.TopMargin),fromsheet.getMargin(Sheet.BottomMargin),fromsheet.getMargin(Sheet.LeftMargin),fromsheet.getMargin(Sheet.RightMargin));
        excel.print_setting_on_sheet_PrintSetup(newsheet,1,0,PrintSetup.LETTER_PAPERSIZE);
        excel.print_setting_on_sheet(newsheet,true,true,true,true,true);

        excel.setPrintArea(excel,0,0,5,0,5);



excel.save();
//excel.outputFile("C:\\Users\\YUY139\\Desktop\\result");
//           PrintSetup ps = newsheet.getPrintSetup();
//            ps.setLandscape(false); //  Printing direction, true: horizontal, false: vertical (default)
//            ps.setVResolution((short)600);
//            ps.setPaperSize(HSSFPrintSetup.A4_PAPERSIZE); //Paper type
//
//            SheetFunc.copyRows(wb, 0, num+1,0 , 46, 0);//copy
//            wb.getSheetAt(num+1).setColumnWidth((short)0, (short)2400);//256,31.38
return excel;
    }}

