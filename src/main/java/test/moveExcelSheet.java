package test;

import msexcel.Excel;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.filechooser.FileSystemView;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import static copyExcel.converter.copySheet;
import static copyExcel.converter.copySheetSettings;

public class moveExcelSheet {

    public static final String template = "template.xlsx";
    public static final String destination = "destination.xlsx";
    public static final String sheetNameToBeDeleted = "delete";
    public static final String SheetNameToBeAdded = "Item 15";

    public static final String template_absolute_path = FileSystemView.getFileSystemView().getHomeDirectory().getPath() + File.separator + template;
    public static final String destination_absolute_path = FileSystemView.getFileSystemView().getHomeDirectory().getPath() + File.separator + destination;

    public static void main(String[] args) throws IOException, InterruptedException, InvalidFormatException

    {
        System.out.println(template_absolute_path);
        System.out.println(destination_absolute_path);
        Workbook wb_destination = new XSSFWorkbook(OPCPackage.open(destination_absolute_path));
        Workbook wb_template = new XSSFWorkbook(OPCPackage.open(template_absolute_path));

//        Excel destination = Excel.loadExcel(destination_absolute_path);
//        Excel template = Excel.loadExcel(template_absolute_path);
//
//        template.assignSheet(wb_template.getSheetIndex(SheetNameToBeAdded));
//        destination.assignSheet(wb_destination.getSheetIndex(sheetNameToBeDeleted));
        System.out.println("sheetToBeDeleted :" + wb_destination.getSheetIndex(sheetNameToBeDeleted) + ", SheetToBeAdded: " + wb_template.getSheetIndex(SheetNameToBeAdded));


        /*Get sheets from the temp file*/
        XSSFSheet sheet_1 = ((XSSFWorkbook) wb_destination).getSheet(sheetNameToBeDeleted);
        XSSFSheet sheet_2 = (XSSFSheet) wb_destination.createSheet(SheetNameToBeAdded);
        XSSFSheet sheet_3 = ((XSSFWorkbook) wb_template).getSheet(SheetNameToBeAdded);

        copySheet(sheet_3, sheet_2);
        copySheetSettings(sheet_3, sheet_2);

//        wb_destination.removeSheetAt(wb_destination.getSheetIndex(sheet_1.getSheetName()));
        //4 新增新表格

        wb_destination.setSheetOrder(SheetNameToBeAdded,2);
        OutputStream os = new FileOutputStream(destination);
        System.out.print("If you arrived here, it means you're good boy");
        wb_destination.write(os);
        os.flush();
        os.close();
        wb_destination.close();
    }
}
