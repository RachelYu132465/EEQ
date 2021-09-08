package validate;

import common.NumberProcessor;
import msexcel.Excel;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import test.ExcelForRu;

import java.util.ArrayList;
import java.util.List;

import static common.StringProcessor.*;

public class excelFormulaProcessor {
    public static final String SPECIFICATION = "specification";
    public static final String ERROR = "error";

//    public static double getTargetOutput(String specification) {
//
//        if (specification.contains("%")) {
//            return Double.valueOf(specification.replaceAll("[^(0-9|.)]", "")) / 100;
//        }
//
//        return Double.valueOf(specification.replaceAll("[^(0-9|.)]", ""));
//    }

    public static String findBlackTitleAtLeft(Excel excel, Cell cell) {

        int rowIdx = cell.getRowIndex();
        int colIdx = cell.getColumnIndex();
        for (int k = colIdx - 1; k >= 0; k--) {
            excel.assignRow(rowIdx);
            excel.assignCell(k);
//            System.out.println("  rowidx:" + k);
            XSSFCellStyle cs = ((XSSFCellStyle) excel.getCurCell().getCellStyle());
            XSSFFont font = cs.getFont();
            XSSFColor color = font.getXSSFColor();
//            System.out.println("Font color : " + color.getARGBHex());
            if ((excel.getCellValue().toString().isEmpty()))
                return "";

            else if (!(excel.getCellValue().toString().isEmpty())) {
                if (excel.getfontColor().equals("NULL") || excel.getfontColor().equals("FF000000")) {
//                    System.out.println("if loop" + k);
                    return excel.getCellValue().toString();

                }
            }
        }
        return "not find title at Left";
    }

    public static String findBlackTitleAtTop(Excel excel, Cell cell) {

        int rowIdx = cell.getRowIndex();
        int colIdx = cell.getColumnIndex();
        for (int k = rowIdx - 1; k >= 0; k--) {
            excel.assignRow(k);
            excel.assignCell(colIdx);

//            System.out.println("  rowidx:" + k);
            // XSSFCellStyle cs = ((XSSFCellStyle)
            // this.getCurCell().getCellStyle());
            // XSSFFont font = cs.getFont();
            // XSSFColor color = font.getXSSFColor();
            // System.out.println("Font color : " + color.getARGBHex());
            if (!(excel.getCellValue().toString().isEmpty())) {
                if (excel.getfontColor().equals("NULL") || excel.getfontColor().equals("FF000000")) {

                    // System.out.println("if loop" + k);
                    return excel.getCellValue().toString();

                }
            }
        }
        return "not find title at Top";
    }

    /*
        開始找字彙 specification
        尋找邏輯: 找下列的該欄，再下列的首欄開始找，再找下下列以此類推
     */
    public static String findSpecification(Excel excel, int currentRowIdx, int currentCellIdx) {
        for (int tgtRowIdx = currentRowIdx + 1; tgtRowIdx < excel.getLastRowNum(); ++tgtRowIdx) {
            excel.assignRow(tgtRowIdx);

            //找下列的該欄
            if (ifContain(SPECIFICATION, excel.assignCell(currentCellIdx).getCellValue().toString())) {
                return findSpecificationValue(excel, currentCellIdx);
            } else {
                //找再下列的首欄
                for (int tgtCelldx = 0; tgtCelldx < excel.getLastCellNum(); ++tgtCelldx) {
                    if (ifContain(SPECIFICATION, excel.assignCell(tgtCelldx).getCellValue().toString())) {
                        return findSpecificationValue(excel, tgtCelldx);
                    }
                }
            }
        }
        //結束找詞彙

        return "找不到規格";
    }

    public static String findSpecificationValue(Excel excel, int currentCellIdx) {
        if (ifEqual(SPECIFICATION, excel.getCellValue().toString())) {
            return excel.findNextCellStrValue();
        } else
            return excel.getCellValue().toString().replace(SPECIFICATION + ":", "");
    }

    public static void main(String[] args) {
        Excel test = Excel.loadExcel(ExcelForRu.proj_path + "/test1.xlsx");
        System.out.println(test.getFile().getAbsolutePath());

    }

    public static String checkFontColor(Workbook wb, Cell cell) {
        CellStyle c = cell.getCellStyle();
        if (!(c == null)) {
            ////For xls (HSSFWorkbook)  or index =12
            if (wb instanceof HSSFWorkbook) {
//                System.out.println("HSSF:" + ((HSSFWorkbook) excel.getWorkbook()).getInternalWorkbook().getFontRecordAt(c.getFontIndexAsInt()).getColorPaletteIndex());

                return "color format in HSSFWorkbook";
            }
            //For xlsx (XSSFWorkbook) rgb="FF0000CC"/> AND index =0
            //or index =12
            if (wb instanceof XSSFWorkbook) {
                XSSFColor color = ((XSSFWorkbook) wb).getFontAt(c.getFontIndexAsInt()).getXSSFColor();
                //arr contains 4 byte --> first one is for index (please ignore)
                if (!(color == null)) {
                    byte[] rgb = color.getARGB();
                    int[] RGB = new int[3];
                    //j 給rgb 四格用 i給RGB
                    for (int i = 0, j = 1; j < rgb.length; i++, j++) {
                        RGB[i] = Math.abs(rgb[j]);
                    }

                    for (int b : RGB) {
                        System.out.print(b + ",");
                    }
                    System.out.println("");
                    int index = NumberProcessor.getIndexOfLargest(RGB);
                    switch (index) {
                        case (0):
//                            System.out.println(excel.getR1C1Idx(cell) + ":" + "red");
                            return "red";
                        case (1):
//                            System.out.println(excel.getR1C1Idx(cell) + ":" + "green");
                            return "green";

                        case (2):
//                            System.out.println(excel.getR1C1Idx(cell) + ":" + "blue");
                            return "blue";
                    }
                }
            }
        }
        return "";
    }

    public static List<Cell> getBlueCells(Workbook wb,Sheet sheet, String color, CellType cellType) {
        List<Cell> result = new ArrayList<>();
        int rowSize = sheet.getLastRowNum();
//        System.out.println(sheet.getSheetName()+"rowSize: "+rowSize);
        for (int rowIdx = 0; rowIdx < rowSize; ++rowIdx) {
            Row row = sheet.getRow(rowIdx);
//            System.out.println(rowIdx);
            if (row != null) {
                int cellsize = row.getLastCellNum();
                for (int cellIndx = 0; cellIndx < cellsize; ++cellIndx) {
                    Cell cell = row.getCell(cellIndx);
                    if (cell != null && checkFontColor(wb, cell).equals(color)) {
                        if (cell.getCellType().equals(cellType)) {
                            result.add(cell);
                        }
                    }
                }
            }
        }
//        }
        return result;
    }

    public static List<String> findNonEmptyValueInMultipleParameter(String parameterSignature) {
        String paramters[] = parameterSignature.split(",");
        List<String> results = new ArrayList<String>();
        for (String parameter : paramters) {
            if (!parameter.replaceAll("\"", "").trim().isBlank()) {
                results.add(parameter);
            }
        }
        return results;
    }

    public static List<String> removeNonFormulaString(List<String> StringListToCheck) {
        List<String> result = new ArrayList<>();
        for (String StringToCheck : StringListToCheck) {
            if (!ifContainMethod(StringToCheck)) {
                result.add(StringToCheck);
            }
        }
        return result;
    }

    public static String findFormulaForValidate(String originalFormula) {
        if (originalFormula.contains(",")) {
            int secondCommaIdx = originalFormula.replaceFirst(",", " ").indexOf(",");
            return originalFormula.substring(
                    secondCommaIdx + 1
                    , originalFormula.length());
        } else return originalFormula;

    }


}
