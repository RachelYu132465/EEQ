package msexcel;


import dataStructure.ValidGoal;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import test.ExcelForRu;
import validate.MyRange;
import validate.OperatorConvertor;




import java.util.HashMap;
import java.util.Map;


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

    public static String[] setConditionalFormattingRule(String cellAddress, MyRange range)
//            throws RangeException
    {


        String ConditionalFormattingRule[] = new String[2];
        String ROUND = "ROUND";
        // String cellAddress = excel.getR1C1Idx();
        String maxDecimalPlace = "";
        String minDecimalPlace = "";

        String maxOperator = "";
        String maxNumber = "";

        String minOperator = "";
        String minNumber = "";

        String roundFuction_max = "";
        String roundFuction_min = "";
//        if ((!range.hasMax()) && (!range.hasMin()))
//            throw new RangeException("range has no max and min");
if (range.getNoRange()==true){
    String Operator = "<>";
    String Number = String.valueOf(range.getMyNumber());
    String DecimalPlace = String.valueOf(range.getMyNumberDecimalPlace());
    String roundFuction = ROUND + "(" + cellAddress + "," + DecimalPlace + ")";
    ConditionalFormattingRule[0] = roundFuction + Operator + Number;
    return ConditionalFormattingRule;
}
        if (range.hasMax()) {
            maxOperator = OperatorConvertor.getUnAcceptableSymbol(false, range.getMaxEqualTo());
            maxNumber = String.valueOf(range.getMax());
            maxDecimalPlace = String.valueOf(range.getMaxDecimalPlace());
            roundFuction_max = ROUND + "(" + cellAddress + "," + maxDecimalPlace + ")";

        }
        if (range.hasMin()) {
            minOperator = OperatorConvertor.getUnAcceptableSymbol(true, range.getMinEqualTo());
            minNumber = String.valueOf(range.getMin());
            minDecimalPlace = String.valueOf(range.getMinDecimalPlace());
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

    public static Excel setMyConditionalFormatting(HashMap<String, ValidGoal> TobeProcessed, Excel excel)
//            throws RangeException
    {
//cellAddress for example:A1
        Sheet sheet = excel.getSheet();
        for (Map.Entry<String, ValidGoal> goals : TobeProcessed.entrySet()) {

            ExcelCell cell = goals.getValue().getOutput();
            String cellAddress = cell.getR1c1();

            String ConditionalFormattingRules[] = new String[2];
//            try {
                ConditionalFormattingRules = setConditionalFormattingRule(cellAddress, goals.getValue().getMyRange());
//            } catch (RangeException e) {
//                e.printStackTrace();
//            }


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
        return excel;
    }

}





