package validate;

import common.NumberProcessor;
import common.StringProcessor;
import org.apache.commons.lang.StringUtils;


import java.math.BigDecimal;
import java.util.*;

public class customStringFormatter {


    //2.傳出Acceptable Range 字串
    public static String getAccecptableRangeString(MyRange myRange) {
        String minRange = "";
        String maxRange = "";
        String andSign = " & ";
        if (myRange.hasMax()) {
            maxRange = OperatorConvertor.getAcceptableSymbol(true, myRange.getMaxEqualTo());
        }

        if (myRange.hasMin()) {
            minRange = myRange.minSymbol.getAcceptableSymbol(false, myRange.getMinEqualTo());
        }

        if (!minRange.isEmpty() && !maxRange.isEmpty()) {
            return minRange + andSign + maxRange;

        } else return minRange + maxRange;

    }

    public static List<String> sperateStringByKeywords(String string, List<String> keywords) {
        int beginIndex = 0;
        int endIndex = 0;
        List<String> subStrings = new ArrayList<>();
        for (int i = 0; i < keywords.size(); i++) {
            String keyword = keywords.get(i);
            endIndex = string.indexOf(keyword);
            String subString = string.substring(beginIndex, endIndex);
            subStrings.add(subString);
            beginIndex = endIndex;

        }
        return subStrings;
    }

    public static List<Double> setOOS (MyRange range ) throws RangeException {
        List<Double> OOSNum =new ArrayList<>();


        // 有上下限
        if (range.hasMax()&&range.hasMin()){


            BigDecimal bd1 = range.getMin();
            //若數字有等於符號，OOS數字: 上限需加1， 下限需減1
            if(range.getMinEqualTo()){
                BigDecimal oneForMin = NumberProcessor.OneInLastDigitPoint(range.getMinDecimalPlace());
                bd1=bd1.subtract(oneForMin);
            }
            OOSNum.add(bd1.doubleValue());

            BigDecimal bd2 = range.getMax();
            if(range.getMaxEqualTo()){
                BigDecimal oneForMax = NumberProcessor.OneInLastDigitPoint(range.getMaxDecimalPlace());
                bd2=bd2.add(oneForMax);
            }
            OOSNum.add(bd2.doubleValue());


        }

//只有上限或下限
        else if (!range.hasMax()&&range.hasMin()){
            BigDecimal bd1 = range.getMin();
            //若數字有等於符號，OOS數字: 上限需加1， 下限需減1
            if(range.getMinEqualTo()){
                BigDecimal oneForMin = NumberProcessor.OneInLastDigitPoint(range.getMinDecimalPlace());
                bd1=bd1.subtract(oneForMin);
            }
            OOSNum.add(bd1.doubleValue());
        }
        else if (range.hasMax()&&!range.hasMin()){

            BigDecimal bd2 = range.getMax();
            if(range.getMaxEqualTo()){
                BigDecimal oneForMax = NumberProcessor.OneInLastDigitPoint(range.getMaxDecimalPlace());
                bd2=bd2.add(oneForMax);
            }
            OOSNum.add(bd2.doubleValue());


        }
        else {
            throw new RangeException("neither max nor min");
        }

        return OOSNum;
    }
    public static MyRange MyRangeGenerator(String specification) {

        MyRange range = new MyRange();
        BigDecimal min;
        BigDecimal max;
        List<String> extractdigitFromString = StringProcessor.extractStringByPattern(specification, StringProcessor.extract_digit_rex);
        List<BigDecimal> extractdigits = new ArrayList<>();
        for (String s : extractdigitFromString) {
            if (StringUtils.isNumeric(s))
                extractdigits.add(new BigDecimal(s));
        }
        List<String> subStrings = sperateStringByKeywords(specification, extractdigitFromString);


        //有上下限，會傳入兩個數字
        if (extractdigitFromString.size() == 2) {


            min = extractdigits.get(0);
            max = extractdigits.get(1);

            Boolean minEqualTo;
            Boolean maxEqualTo;

            minEqualTo = OperatorConvertor.isEqual(subStrings.get(0));
            maxEqualTo = OperatorConvertor.isEqual(subStrings.get(1));

            range.setMinRange(min, minEqualTo);
            range.setMaxRange(max, maxEqualTo);

            int minDecimalplace = NumberProcessor.countDecimalPlace(extractdigitFromString.get(0));
            int maxDecimalplace = NumberProcessor.countDecimalPlace(extractdigitFromString.get(1));


            range.setMinDecimalPlace(minDecimalplace);
            range.setMaxDecimalPlace(maxDecimalplace);
        }
        //只有上下限其中一個，會傳入一個數字
        if (extractdigitFromString.size() == 1) {

            //判斷是大於或小於
            if (OperatorConvertor.isGreater(subStrings.get(0))) {

                max = extractdigits.get(0);
                Boolean maxEqualTo;
                maxEqualTo = OperatorConvertor.isEqual(subStrings.get(0));
                range.setMaxRange(max, maxEqualTo);

            }
            else {
                min = extractdigits.get(0);
                Boolean minEqualTo;
                minEqualTo = OperatorConvertor.isEqual(subStrings.get(0));
                range.setMinRange(min, minEqualTo);
            }
        }

        return range;
    }



    public static void main(String[] args) {
        MyRange myRange =MyRangeGenerator("not more than 5%");
        System.out.println(myRange.toString());
    }

}
