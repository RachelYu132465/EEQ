package msexcel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import test.ExcelForRu;
import validate.MyComparision;

import static common.NumberProcessor.countDecimalPlace;

public class customExcelStyle {
    public static void main(String[] args) {
        Excel test = Excel.loadExcel(ExcelForRu.proj_path + "/testR.xlsx");
        Cell c =test.assignSheet(0).assignRow(0).assignCell(0).getCurCell();
        System.out.println(test.getCellValue() + c.getCellType().name() );
//        MyComparision m = new MyComparision(" Not more than 3.0% is retained on No.80 U.S. standard sieve.");

        MyComparision m = new MyComparision("The Recovery is not less than 97.5% and not  more than 102.5%.");
        String[] s = setConditionalFormattingRule(test,m);
        for (String ss : s){
            if(!(ss==null))
            setMyConditionalFormatting(test.getSheet(),ss);
        }
//        test.saveToFile();
        test.outputFile("C:\\Users\\YUY139\\Desktop\\result");
    }
    public static String[] setConditionalFormattingRule(Excel excel, MyComparision spec){
        String ConditionalFormattingRule[]=new String [2];
        if (spec.getTargetSmall()&& spec.getTragetBig()){

          int decimalNum1 = countDecimalPlace(spec.getNumS()[0] );
            ConditionalFormattingRule[0] = "ROUND("+ excel.getR1C1Idx()+","+decimalNum1  +")<" + spec.getNum()[0];

            int decimalNum2 = countDecimalPlace(spec.getNumS()[1] );
            ConditionalFormattingRule[1] = "ROUND("+ excel.getR1C1Idx() +","+decimalNum2 +")>" + spec.getNum()[1];
        }

else if (spec.getTragetBig()&& !spec.getTargetSmall()){
            int decimalNum1 = countDecimalPlace(spec.getNumS()[0] );
            ConditionalFormattingRule[0] = "ROUND("+ excel.getR1C1Idx()+","+decimalNum1  +")<" + spec.getNum()[0];

        }
else if (spec.getTargetSmall()&& !spec.getTragetBig()){
            int decimalNum2 = countDecimalPlace(spec.getNumS()[0] );
            ConditionalFormattingRule[0] = "ROUND("+ excel.getR1C1Idx() +","+decimalNum2 +")>" + spec.getNum()[0];
        }
        System.out.println(ConditionalFormattingRule[0]);
        return ConditionalFormattingRule;
    }

    public static void setMyConditionalFormatting(Sheet sheet,String ConditionalFormattingRule){

        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        ConditionalFormattingRule rule = sheetCF.createConditionalFormattingRule(ConditionalFormattingRule);
        PatternFormatting fill = rule.createPatternFormatting();

        fill.setFillBackgroundColor(IndexedColors.RED1.index);
        FontFormatting fontFmt = rule.createFontFormatting();

        fontFmt.setFontColorIndex(IndexedColors.WHITE.index);

        ConditionalFormattingRule[] cfRules = new ConditionalFormattingRule[]{rule};

        CellRangeAddress[] regions = new CellRangeAddress[]{CellRangeAddress.valueOf("A1")};

        sheetCF.addConditionalFormatting(regions, cfRules);
    }
}
