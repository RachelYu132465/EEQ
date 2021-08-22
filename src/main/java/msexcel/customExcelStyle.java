package msexcel;

import common.NumberProcessor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import test.ExcelForRu;
import validate.MyRange;
import validate.OperatorConvertor;
import validate.RangeException;


import java.math.BigDecimal;

import static common.NumberProcessor.countDecimalPlace;

public class customExcelStyle {
    public static void main(String[] args) {
        Excel test = Excel.loadExcel(ExcelForRu.proj_path + "/testR.xlsx");
        Cell c = test.assignSheet(0).assignRow(0).assignCell(0).getCurCell();
        System.out.println(test.getCellValue() + c.getCellType().name());
//        MyComparision m = new MyComparision(" Not more than 3.0% is retained on No.80 U.S. standard sieve.");

//        MyComparision m = new MyComparision("The Recovery is not less than 97.5% and not  more than 102.5%.");
//        String[] s = setConditionalFormattingRule(test,m);
//
//                setMyConditionalFormatting(test.getSheet(), ss);
//
//        test.saveToFile();
        test.outputFile("C:\\Users\\YUY139\\Desktop\\result");
    }

    public static String[] setConditionalFormattingRule(Excel excel, MyRange range) throws RangeException {


        String ConditionalFormattingRule[] = new String[2];
        String ROUND = "ROUND";
        String cellAddress = excel.getR1C1Idx();
        String maxDecimalPlace = "";
        String minDecimalPlace = "";

        String maxOperator = "";
        String maxNumber = "";

        String minOperator = "";
        String minNumber = "";

        String roundFuction_max = "";
        String roundFuction_min = "";
        if ((!range.hasMax()) && (!range.hasMin()))
            throw new RangeException("range has no max and min");

        if (range.hasMax()) {
            maxOperator = OperatorConvertor.getUnAcceptableSymbol(true, range.getMaxEqualTo());
            maxNumber = String.valueOf(range.getMax());
            maxDecimalPlace = String.valueOf(maxDecimalPlace);
            roundFuction_max = ROUND + "(" + cellAddress + "," + maxDecimalPlace + ")";

        }
        if (range.hasMin()) {
            minOperator = OperatorConvertor.getUnAcceptableSymbol(false, range.getMinEqualTo());
            minNumber = String.valueOf(range.getMin());
            minDecimalPlace = String.valueOf(minDecimalPlace);
            roundFuction_min = ROUND + "(" + cellAddress + "," + minDecimalPlace + ")";

        }
        if (range.hasMax() && range.hasMin()) {
            ConditionalFormattingRule[0] = roundFuction_max + maxOperator + maxNumber;
            ConditionalFormattingRule[1] = roundFuction_min + minOperator + minNumber;
        } else if (!range.hasMax() && range.hasMin()) {

            ConditionalFormattingRule[0] = roundFuction_min + minOperator + minNumber;
        } else if (range.hasMax() && !range.hasMin()) {

            ConditionalFormattingRule[0] = roundFuction_max + maxOperator + maxNumber;

        }

        return ConditionalFormattingRule;
    }

    public static void setMyConditionalFormatting(String cellAddress, Sheet sheet, String[] ConditionalFormattingRules) {
//cellAddress for example:A1
        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        for (String ConditionalFormattingRule : ConditionalFormattingRules) {
            ConditionalFormattingRule rule = sheetCF.createConditionalFormattingRule(ConditionalFormattingRule);
            PatternFormatting fill = rule.createPatternFormatting();

            fill.setFillBackgroundColor(IndexedColors.RED1.index);
            FontFormatting fontFmt = rule.createFontFormatting();

            fontFmt.setFontColorIndex(IndexedColors.WHITE.index);

            ConditionalFormattingRule[] cfRules = new ConditionalFormattingRule[]{rule};

            CellRangeAddress[] regions = new CellRangeAddress[]{CellRangeAddress.valueOf(cellAddress)};

            sheetCF.addConditionalFormatting(regions, cfRules);
        }

    }
}
