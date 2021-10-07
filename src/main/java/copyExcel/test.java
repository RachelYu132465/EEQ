package copyExcel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Arrays;

import msexcel.Excel;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.filechooser.FileSystemView;

import static msexcel.Excel.ExcelType.EXCEL_XLSX;

public class test {
    public static final String template = "Bulklist.xlsx";
    public static final String template_absolute_path = FileSystemView.getFileSystemView().getHomeDirectory().getPath() + File.separator + template;

    public static void main(String[] args) {

            Excel excel2  =Excel.createExcel(EXCEL_XLSX);
        Excel excel = Excel.loadExcel(template_absolute_path);


//            int sheetSize =excel.getNumberOfSheets();
//            for (int sheetCnt = 0; sheetCnt < sheetSize; ++sheetCnt) {
        excel.assignSheet(0);

//                excel.assignSheet(sheetCnt);
//                Sheet sheet = excel.assignSheet(sheetCnt).getSheet();
        int rowSize = excel.getLastRowNum();
        for (int j = 1; j < rowSize; j++) {
            String fileName = "";
            excel.assignRow(j);
            int colSize = excel.getLastCellNum();
            for (int k = 0; k < 2; k++) {
                excel.assignCell(k);
String cell = excel.getStringCellValue();
                fileName = fileName + cell;
                if (k == 0) {
                    fileName = fileName + "-";
//
                }

                if (k == 1) {
                    if(StringUtils.isBlank(excel.getStringCellValue()) ){
                        StringUtils.remove(fileName, "-");
                    }
                        try (OutputStream os = new FileOutputStream(fileName + ".xlsx")) {
                            Workbook wb = new XSSFWorkbook();
//                            Sheet sheet1 =
                                    wb.createSheet("sheet");
//                            Sheet sheet2 =
                                    wb.createSheet("STB");

                            wb.write(os);
                        } catch (Exception e) {
                            System.out.println(e.getMessage());
                        }



                }
            }
        }
    }
//            try (OutputStream os = new FileOutputStream("Javatpoint.xls")) {
//                Workbook wb = new HSSFWorkbook();
//                Sheet sheet = wb.createSheet("Sheet");
//                Row row     = sheet.createRow(1);
//                Row row2    = sheet.createRow(2);
//                Cell cell   = row.createCell(1);
//                Cell cell2  = row2.createCell(1);
//                cell.setCellValue("Hello");
//                cell2.setCellValue("Hello, Javatpoint");
//                sheet.shiftRows(1, 2, 1);
//                wb.write(os);
//            }catch(Exception e) {
//                System.out.println(e.getMessage());
//            }
//        }
}

